
/* =========================================================
   ALFERZA LIVE OFFICE — INFRAESTRUCTURA TI
   Microsoft Graph + SharePoint
   Inventario, monitoreo y gestión de trabajadores
========================================================= */

"use strict";

/* ========================= CONFIGURACIÓN ========================= */

const CONFIG = {
    clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
    tenantId: "dbab984f-4bb1-4b60-9dff-da59f54acdf1",
    redirectUri: "https://alferzati.github.io/live-office/blank.html",
    hostname: "alferzaholding-my.sharepoint.com",
    sitePath: "/personal/soporte1_alferza_pe",
    listaServicios: "MonitoreoTI",
    listaTrabajadores: "TrabajadoresEquipos",
    archivoEquipos: "Procedimientos T.I/equipos.json",
    driveId: "b!1l0EwOKXpU24cUGYaNjPRiTQWIyIIChCkDDL197LcoOFOD7ThItqTIJ_Z1KoYrrS",
    scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"],
    cachePrefix: "alferza_ti_",
    equiposPorPagina: 40,
    intervalo: 120000
};

const msalApp = typeof msal !== "undefined"
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

/* ========================= ESTADO ========================= */

let cuentaActual = null;
let tokenActual = null;
let sitio = null;
let listaTrabajadores = null;
let columnasTrabajadores = null;
let promesaSitio = null;
let promesaLista = null;
let promesaColumnas = null;
let promesaTrabajadores = null;

let equiposData = [];
let equiposFiltrados = [];
let trabajadoresData = [];
let trabajadoresPorEquipo = new Map();

let paginaActual = 1;
let equipoGestionSeleccionado = null;
let asignacionGestionSeleccionada = null;
let ubicacionEquipos = null;
let etagEquipos = null;
let cargandoEquipos = false;
let trabajadoresCargados = false;
let ultimaActualizacion = 0;
let temporizadorBusqueda = null;

/* ========================= ELEMENTOS HTML ========================= */

const $ = id => document.getElementById(id);

const el = {
    loading: $("loadingOverlay"),
    loadingMessage: $("loadingMessage"),

    serviciosGrid: $("serviciosGrid"),
    estadoGeneral: $("estadoGeneral"),

    tabla: $("equiposTableBody"),
    vacio: $("equiposEmpty"),
    resultados: $("equiposResultados"),

    total: $("totalEquipos"),
    conectados: $("equiposConectados"),
    desconectados: $("equiposDesconectados"),

    filtroEstado: $("filtroEstadoEquipo"),
    filtroArea: $("filtroAreaEquipo"),
    buscar: $("buscarEquipo"),
    estadoInventario: $("equiposEstado"),
    ultimaRevision: $("lastUpdate"),

    modal: $("gestionTrabajadorModal"),
    ip: $("gestionIP"),
    buscarIP: $("buscarEquipoIP"),
    equipoEncontrado: $("equipoEncontrado"),
    equipoNombre: $("gestionEquipoNombre"),
    equipoDetalle: $("gestionEquipoDetalle"),
    trabajador: $("gestionTrabajador"),
    mensaje: $("gestionTrabajadorMensaje"),

    abrirGestion: $("btnGestionarTrabajadores"),
    registrar: $("btnRegistrarTrabajador"),
    editar: $("btnEditarTrabajador"),
    eliminar: $("btnEliminarTrabajador"),
    cerrar: $("cerrarGestionTrabajadorModal"),
    cancelar: $("cancelarGestionTrabajador")
};

/* ========================= UTILIDADES ========================= */

function normalizar(valor) {
    return String(valor ?? "").trim().toLowerCase();
}

function normalizarClave(valor) {
    return normalizar(valor)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-z0-9]/g, "");
}

function escaparHTML(valor) {
    return String(valor ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function fechaActual() {
    return new Date().toLocaleString("es-PE", {
        dateStyle: "short",
        timeStyle: "medium"
    });
}

function cacheLeer(clave) {
    try {
        const dato = localStorage.getItem(CONFIG.cachePrefix + clave);
        return dato ? JSON.parse(dato) : null;
    } catch {
        return null;
    }
}

function cacheGuardar(clave, dato) {
    try {
        localStorage.setItem(
            CONFIG.cachePrefix + clave,
            JSON.stringify(dato)
        );
    } catch {
        // La aplicación sigue funcionando sin caché.
    }
}

function mostrarMensaje(texto, tipo = "error") {
    if (!el.mensaje) return;

    el.mensaje.textContent = texto;
    el.mensaje.style.display = "block";
    el.mensaje.className = `gestion-mensaje ${tipo}`;
}

function limpiarMensaje() {
    if (!el.mensaje) return;

    el.mensaje.textContent = "";
    el.mensaje.style.display = "none";
}

function mostrarError(error, operacion) {
    console.error(`[ALFERZA TI] ${operacion}:`, error);

    const mensaje = error?.message || "Error inesperado.";

    mostrarMensaje(`${operacion}: ${mensaje}`, "error");
}

function mostrarCarga() {
    if (el.loading) el.loading.style.display = "flex";
}

function ocultarCarga() {
    if (el.loading) el.loading.style.display = "none";
}

/* ========================= AUTENTICACIÓN ========================= */

async function obtenerTokenSilencioso() {
    if (!msalApp) return null;

    const cuentas = msalApp.getAllAccounts();

    if (!cuentas.length) return null;

    cuentaActual = cuentas[0];

    try {
        const resultado = await msalApp.acquireTokenSilent({
            scopes: CONFIG.scopes,
            account: cuentaActual
        });

        tokenActual = resultado.accessToken;
        return tokenActual;
    } catch (error) {
        console.warn("[ALFERZA TI] Token silencioso no disponible.", error);
        return null;
    }
}

async function iniciarSesion(interactivo = false) {
    const token = await obtenerTokenSilencioso();

    if (token || !interactivo) return token;

    if (!msalApp) {
        throw new Error("La biblioteca MSAL no está cargada.");
    }

    const resultado = cuentaActual
        ? await msalApp.acquireTokenPopup({
            scopes: CONFIG.scopes,
            account: cuentaActual
        })
        : await msalApp.loginPopup({
            scopes: CONFIG.scopes
        });

    cuentaActual = resultado.account;
    tokenActual = resultado.accessToken ||
        await obtenerTokenSilencioso();

    return tokenActual;
}

/* ========================= MICROSOFT GRAPH ========================= */

async function graphFetch(url, opciones = {}, intento = 0) {
    const { timeoutMs = 30000, ...configuracion } = opciones;

    if (!tokenActual) {
        const token = await obtenerTokenSilencioso();

        if (!token) {
            throw new Error("Sesión de Microsoft no disponible.");
        }
    }

    const controlador = new AbortController();
    const temporizador = setTimeout(
        () => controlador.abort(),
        timeoutMs
    );

    try {
        const respuesta = await fetch(url, {
            ...configuracion,
            signal: controlador.signal,
            headers: {
                Authorization: `Bearer ${tokenActual}`,
                ...(configuracion.headers || {})
            }
        });

        if (
            respuesta.status === 401 &&
            intento < 1 &&
            await obtenerTokenSilencioso()
        ) {
            return graphFetch(url, opciones, intento + 1);
        }

        if (
            [429, 503].includes(respuesta.status) &&
            intento < 3
        ) {
            const segundos = Number(
                respuesta.headers.get("Retry-After")
            ) || Math.pow(2, intento);

            await new Promise(resolve =>
                setTimeout(resolve, Math.min(segundos * 1000, 10000))
            );

            return graphFetch(url, opciones, intento + 1);
        }

        if (!respuesta.ok) {
            let detalle = "";

            try {
                const json = await respuesta.json();
                detalle = json.error?.message || "";
            } catch {
                // Puede no existir cuerpo JSON.
            }

            const error = new Error(
                detalle ||
                `Microsoft Graph respondió HTTP ${respuesta.status}.`
            );

            error.status = respuesta.status;
            throw error;
        }

        return respuesta;
    } catch (error) {
        if (error.name === "AbortError") {
            throw new Error("La solicitud a SharePoint agotó el tiempo de espera.");
        }

        throw error;
    } finally {
        clearTimeout(temporizador);
    }
}

async function graphJSON(url, opciones) {
    const respuesta = await graphFetch(url, opciones);

    if (respuesta.status === 204) return null;

    return respuesta.json();
}

async function obtenerTodosLosElementos(url) {
    const elementos = [];

    while (url) {
        const datos = await graphJSON(url);
        elementos.push(...(datos.value || []));
        url = datos["@odata.nextLink"] || null;
    }

    return elementos;
}

/* ========================= SITIO Y LISTAS SHAREPOINT ========================= */

function obtenerSitio() {
    if (!promesaSitio) {
        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${CONFIG.hostname}:${CONFIG.sitePath}?$select=id`;

        promesaSitio = graphJSON(url)
            .then(resultado => {
                sitio = resultado;
                return sitio;
            })
            .catch(error => {
                promesaSitio = null;
                throw error;
            });
    }

    return promesaSitio;
}

function obtenerListaTrabajadores() {
    if (!promesaLista) {
        promesaLista = (async () => {
            const sitioActual = await obtenerSitio();

            const listas = await obtenerTodosLosElementos(
                `https://graph.microsoft.com/v1.0/sites/${sitioActual.id}/lists?$select=id,displayName,name`
            );

            const lista = listas.find(item =>
                normalizar(item.displayName) ===
                    normalizar(CONFIG.listaTrabajadores) ||
                normalizar(item.name) ===
                    normalizar(CONFIG.listaTrabajadores)
            );

            if (!lista) {
                throw new Error(
                    `No se encontró la lista "${CONFIG.listaTrabajadores}".`
                );
            }

            listaTrabajadores = lista;
            return lista;
        })().catch(error => {
            promesaLista = null;
            throw error;
        });
    }

    return promesaLista;
}

function obtenerColumnasTrabajadores() {
    if (!promesaColumnas) {
        promesaColumnas = (async () => {
            const sitioActual = await obtenerSitio();
            const lista = await obtenerListaTrabajadores();

            const columnas = await obtenerTodosLosElementos(
                `https://graph.microsoft.com/v1.0/sites/${sitioActual.id}/lists/${lista.id}/columns?$select=name,displayName,hidden,readOnly`
            );

            columnasTrabajadores = columnas;
            return columnas;
        })().catch(error => {
            promesaColumnas = null;
            throw error;
        });
    }

    return promesaColumnas;
}

/*
 * Obtiene el nombre interno de una columna usando su nombre visible.
 * Esto evita depender de que SharePoint use exactamente el mismo
 * nombre interno y visible.
 */
function buscarColumna(columnas, alternativas) {
    const buscadas = alternativas.map(normalizarClave);

    return columnas.find(columna =>
        buscadas.includes(normalizarClave(columna.displayName))
    ) || columnas.find(columna =>
        buscadas.includes(normalizarClave(columna.name))
    ) || null;
}


async function obtenerEsquemaTrabajadores() {
    const columnas = await obtenerColumnasTrabajadores();

    const columnaEquipo = buscarColumna(columnas, [
        "Equipo1",
        "Equipo",
        "NombreEquipo"
    ]);

    const columnaTrabajador = buscarColumna(columnas, [
        "NombreTrabajador",
        "Trabajador",
        "Nombre del trabajador"
    ]);

    if (!columnaEquipo || !columnaTrabajador) {
        throw new Error(
            "No se encontraron las columnas necesarias en " +
            "TrabajadoresEquipos. Columnas disponibles: " +
            columnas.map(c => `${c.displayName} (${c.name})`).join(", ")
        );
    }

    return {
        equipo: columnaEquipo.name,
        trabajador: columnaTrabajador.name
    };
}


/* ========================= INVENTARIO EQUIPOS.JSON ========================= */

async function buscarArchivoEquipos(raiz) {
    const nombre = CONFIG.archivoEquipos.split("/").pop();

    const respuesta = await graphFetch(
        `${raiz}/root/search(q='${encodeURIComponent(nombre)}')?$select=id,name,eTag&$top=100`
    );

    const datos = await respuesta.json();

    const archivo = (datos.value || []).find(item =>
        normalizar(item.name) === normalizar(nombre)
    );

    if (!archivo) {
        throw new Error(
            `No se encontró ${CONFIG.archivoEquipos} en SharePoint.`
        );
    }

    return {
        meta: `${raiz}/items/${archivo.id}`,
        contenido: `${raiz}/items/${archivo.id}/content`,
        etag: archivo.eTag || null
    };
}

function interpretarJSONEquipos(buffer) {
    const bytes = new Uint8Array(buffer);
    let encoding = "utf-8";

    if (bytes[0] === 0xFF && bytes[1] === 0xFE) {
        encoding = "utf-16le";
    } else if (bytes[0] === 0xFE && bytes[1] === 0xFF) {
        encoding = "utf-16be";
    }

    const texto = new TextDecoder(encoding)
        .decode(buffer)
        .replace(/^\uFEFF/, "");

    let datos;

    try {
        datos = JSON.parse(texto);
    } catch {
        throw new Error("El archivo equipos.json contiene JSON inválido.");
    }

    if (Array.isArray(datos)) return datos;

    if (datos && typeof datos === "object") {
        for (const clave of ["value", "equipos", "Equipos", "data", "items"]) {
            if (Array.isArray(datos[clave])) return datos[clave];
        }

        if (datos.Nombre) return [datos];
    }

    throw new Error("El formato de equipos.json no es válido.");
}

async function obtenerEquiposSharePoint() {
    const raiz =
        `https://graph.microsoft.com/v1.0/drives/${CONFIG.driveId}`;

    const ruta = CONFIG.archivoEquipos
        .split("/")
        .map(encodeURIComponent)
        .join("/");

    let ubicacion = ubicacionEquipos;

    if (!ubicacion) {
        const url =
            `${raiz}/root:/${ruta}?$select=id,name,eTag`;

        try {
            const metadata = await graphJSON(url);

            ubicacion = {
                meta: `${raiz}/items/${metadata.id}`,
                contenido: `${raiz}/items/${metadata.id}/content`,
                etag: metadata.eTag || null
            };
        } catch (error) {
            if (error.status !== 404) throw error;
            ubicacion = await buscarArchivoEquipos(raiz);
        }
    } else {
        const metadata = await graphJSON(
            `${ubicacion.meta}?$select=eTag`
        );

        ubicacion.etag = metadata.eTag || null;
    }

    if (
        ubicacion.etag &&
        ubicacion.etag === etagEquipos &&
        equiposData.length > 0
    ) {
        ubicacionEquipos = ubicacion;
        return null;
    }

    const respuesta = await graphFetch(ubicacion.contenido, {
        timeoutMs: 90000
    });

    const equipos = interpretarJSONEquipos(
        await respuesta.arrayBuffer()
    );

    ubicacionEquipos = ubicacion;
    etagEquipos = ubicacion.etag;

    return equipos;
}

/* ========================= TRABAJADORES: CARGA Y ACTUALIZACIÓN ========================= */

function obtenerAsignacion(nombreEquipo) {
    return trabajadoresPorEquipo.get(normalizar(nombreEquipo)) || null;
}

async function obtenerTrabajadores(forzar = false) {
    if (trabajadoresCargados && !forzar) {
        return trabajadoresPorEquipo;
    }

    if (promesaTrabajadores) return promesaTrabajadores;

    promesaTrabajadores = (async () => {
        const sitioActual = await obtenerSitio();
        const lista = await obtenerListaTrabajadores();

        const elementos = await obtenerTodosLosElementos(
            `https://graph.microsoft.com/v1.0/sites/${sitioActual.id}/lists/${lista.id}/items?expand=fields&$top=200`
        );

        const esquema = await obtenerEsquemaTrabajadores();

        trabajadoresData = elementos;
        trabajadoresPorEquipo = new Map();

        for (const elemento of elementos) {
            const campos = elemento.fields || {};
            const equipo = campos[esquema.equipo];
            const trabajador = campos[esquema.trabajador];

            if (!equipo || !trabajador) continue;

            const clave = normalizar(equipo);

            /*
             * Si hay duplicados, se conservan en trabajadoresData,
             * pero se muestra la primera asignación por equipo.
             */
            if (!trabajadoresPorEquipo.has(clave)) {
                trabajadoresPorEquipo.set(clave, {
                    nombre: trabajador,
                    itemId: String(elemento.id),
                    equipo: String(equipo)
                });
            }
        }

        trabajadoresCargados = true;

        renderizarTablaEquipos();
        return trabajadoresPorEquipo;
    })().catch(error => {
        trabajadoresCargados = false;
        throw error;
    }).finally(() => {
        promesaTrabajadores = null;
    });

    return promesaTrabajadores;
}

function actualizarAsignacionLocal(equipo, nombre, itemId) {
    trabajadoresPorEquipo.set(normalizar(equipo), {
        nombre,
        itemId: String(itemId),
        equipo
    });

    renderizarTablaEquipos();
}

function quitarAsignacionLocal(equipo) {
    trabajadoresPorEquipo.delete(normalizar(equipo));
    renderizarTablaEquipos();
}

/* ========================= CRUD DE TRABAJADORES ========================= */

async function crearAsignacion(equipo, nombre) {
    const sitioActual = await obtenerSitio();
    const lista = await obtenerListaTrabajadores();
    const esquema = await obtenerEsquemaTrabajadores();

    const fields = {};
    fields[esquema.equipo] = equipo;
    fields[esquema.trabajador] = nombre;

    const resultado = await graphJSON(
        `https://graph.microsoft.com/v1.0/sites/${sitioActual.id}/lists/${lista.id}/items`,
        {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ fields })
        }
    );

    return resultado;
}

async function actualizarAsignacion(itemId, nombre) {
    const sitioActual = await obtenerSitio();
    const lista = await obtenerListaTrabajadores();
    const esquema = await obtenerEsquemaTrabajadores();

    const fields = {};
    fields[esquema.trabajador] = nombre;

    return graphJSON(
        `https://graph.microsoft.com/v1.0/sites/${sitioActual.id}/lists/${lista.id}/items/${encodeURIComponent(itemId)}/fields`,
        {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(fields)
        }
    );
}

async function borrarAsignacion(itemId) {
    const sitioActual = await obtenerSitio();
    const lista = await obtenerListaTrabajadores();

    await graphFetch(
        `https://graph.microsoft.com/v1.0/sites/${sitioActual.id}/lists/${lista.id}/items/${encodeURIComponent(itemId)}`,
        { method: "DELETE" }
    );
}

/* ========================= GESTIÓN DE TRABAJADORES ========================= */

function establecerEstadoBotones({ registrar, editar, eliminar }) {
    if (el.registrar) el.registrar.disabled = !registrar;
    if (el.editar) el.editar.disabled = !editar;
    if (el.eliminar) el.eliminar.disabled = !eliminar;
}

function establecerBotonOcupado(boton, ocupado, textoOcupado, textoNormal) {
    if (!boton) return;

    boton.disabled = ocupado;
    boton.textContent = ocupado ? textoOcupado : textoNormal;
}

function limpiarGestion() {
    equipoGestionSeleccionado = null;
    asignacionGestionSeleccionada = null;

    if (el.ip) el.ip.value = "";
    if (el.trabajador) {
        el.trabajador.value = "";
        el.trabajador.disabled = true;
    }

    if (el.equipoEncontrado) {
        el.equipoEncontrado.style.display = "none";
    }

    limpiarMensaje();
    establecerEstadoBotones({
        registrar: false,
        editar: false,
        eliminar: false
    });
}

function abrirGestionTrabajadores() {
    if (!el.modal) return;

    limpiarGestion();
    el.modal.style.display = "flex";

    setTimeout(() => el.ip?.focus(), 100);
}

function cerrarGestionTrabajadores() {
    if (el.modal) el.modal.style.display = "none";
    limpiarGestion();
}

async function buscarEquipoPorIP() {
    const ip = el.ip?.value.trim();

    limpiarMensaje();

    if (!ip) {
        mostrarMensaje("Ingresa una dirección IP.", "error");
        el.ip?.focus();
        return;
    }

    equipoGestionSeleccionado = null;
    asignacionGestionSeleccionada = null;

    if (el.equipoEncontrado) {
        el.equipoEncontrado.style.display = "none";
    }

    if (el.trabajador) {
        el.trabajador.value = "";
        el.trabajador.disabled = true;
    }

    establecerEstadoBotones({
        registrar: false,
        editar: false,
        eliminar: false
    });

    establecerBotonOcupado(
        el.buscarIP,
        true,
        "Buscando...",
        "Buscar"
    );

    try {
        if (!equiposData.length) {
            await cargarEquipos();
        }

        const equipo = equiposData.find(item =>
            normalizar(item.IP) === normalizar(ip)
        );

        if (!equipo) {
            mostrarMensaje(
                `No se encontró ningún equipo con la IP ${ip}.`,
                "error"
            );
            return;
        }

        /*
         * Se carga siempre desde SharePoint antes de decidir si
         * registrar, editar o eliminar.
         */
        await obtenerTrabajadores(true);

        equipoGestionSeleccionado = equipo;

        if (el.equipoNombre) {
            el.equipoNombre.textContent = equipo.Nombre || "Sin nombre";
        }

        if (el.equipoDetalle) {
            el.equipoDetalle.textContent = [
                equipo.IP || ip,
                equipo.DNSHostName || "Sin DNS",
                equipo.Area || "Sin área"
            ].join(" · ");
        }

        if (el.equipoEncontrado) {
            el.equipoEncontrado.style.display = "flex";
        }

        const asignacion = obtenerAsignacion(equipo.Nombre);
        asignacionGestionSeleccionada = asignacion;

        if (el.trabajador) {
            el.trabajador.disabled = false;
            el.trabajador.value = asignacion?.nombre || "";
        }

        if (asignacion) {
            establecerEstadoBotones({
                registrar: false,
                editar: true,
                eliminar: true
            });
        } else {
            establecerEstadoBotones({
                registrar: true,
                editar: false,
                eliminar: false
            });
        }

        el.trabajador?.focus();
    } catch (error) {
        mostrarError(error, "No se pudo buscar el equipo");
    } finally {
        establecerBotonOcupado(
            el.buscarIP,
            false,
            "Buscando...",
            "Buscar"
        );
    }
}

async function registrarTrabajadorPorIP() {
    const equipo = equipoGestionSeleccionado;
    const nombre = el.trabajador?.value.trim();

    limpiarMensaje();

    if (!equipo) {
        mostrarMensaje("Primero busca un equipo por IP.", "error");
        return;
    }

    if (!nombre) {
        mostrarMensaje("Ingresa el nombre del trabajador.", "error");
        el.trabajador?.focus();
        return;
    }

    establecerBotonOcupado(
        el.registrar,
        true,
        "Registrando...",
        "Registrar"
    );

    try {
        await obtenerTrabajadores(true);

        if (obtenerAsignacion(equipo.Nombre)) {
            throw new Error(
                "Este equipo ya tiene un trabajador asignado. " +
                "Vuelve a buscarlo para editar la asignación."
            );
        }

        const resultado = await crearAsignacion(equipo.Nombre, nombre);

        if (!resultado?.id) {
            throw new Error(
                "SharePoint no devolvió el ID de la asignación creada."
            );
        }

        actualizarAsignacionLocal(
            equipo.Nombre,
            nombre,
            resultado.id
        );

        asignacionGestionSeleccionada = obtenerAsignacion(equipo.Nombre);

        el.trabajador.value = nombre;

        establecerEstadoBotones({
            registrar: false,
            editar: true,
            eliminar: true
        });

        mostrarMensaje(
            "Trabajador registrado correctamente.",
            "success"
        );
    } catch (error) {
        mostrarError(error, "No se pudo registrar el trabajador");
    } finally {
        if (equipoGestionSeleccionado === equipo) {
            el.registrar.disabled = !!obtenerAsignacion(equipo.Nombre);
        }
    }
}

async function editarTrabajadorPorIP() {
    const equipo = equipoGestionSeleccionado;
    const nombre = el.trabajador?.value.trim();

    limpiarMensaje();

    if (!equipo) {
        mostrarMensaje("Primero busca un equipo por IP.", "error");
        return;
    }

    if (!nombre) {
        mostrarMensaje("Ingresa el nombre del trabajador.", "error");
        el.trabajador?.focus();
        return;
    }

    establecerBotonOcupado(
        el.editar,
        true,
        "Guardando...",
        "Editar"
    );

    try {
        /*
         * Refresca la lista para evitar modificar un registro
         * que haya sido cambiado por otro usuario.
         */
        await obtenerTrabajadores(true);

        const asignacion = obtenerAsignacion(equipo.Nombre);

        if (!asignacion?.itemId) {
            throw new Error(
                "No se encontró el ID de la asignación. " +
                "Vuelve a buscar el equipo."
            );
        }

        if (asignacion.nombre === nombre) {
            mostrarMensaje("El nombre no ha cambiado.", "success");
            return;
        }

        await actualizarAsignacion(asignacion.itemId, nombre);

        actualizarAsignacionLocal(
            equipo.Nombre,
            nombre,
            asignacion.itemId
        );

        asignacionGestionSeleccionada = obtenerAsignacion(equipo.Nombre);
        el.trabajador.value = nombre;

        mostrarMensaje(
            "Asignación actualizada correctamente.",
            "success"
        );
    } catch (error) {
        mostrarError(error, "No se pudo editar el trabajador");
    } finally {
        establecerBotonOcupado(
            el.editar,
            false,
            "Guardando...",
            "Editar"
        );

        if (equipoGestionSeleccionado) {
            establecerEstadoBotones({
                registrar: false,
                editar: true,
                eliminar: true
            });
        }
    }
}

async function eliminarTrabajadorPorIP() {
    const equipo = equipoGestionSeleccionado;

    limpiarMensaje();

    if (!equipo) {
        mostrarMensaje("Primero busca un equipo por IP.", "error");
        return;
    }

    await obtenerTrabajadores(true).catch(error => {
        mostrarError(error, "No se pudo comprobar la asignación");
        throw error;
    });

    const asignacion = obtenerAsignacion(equipo.Nombre);

    if (!asignacion?.itemId) {
        mostrarMensaje(
            "No se encontró una asignación para este equipo.",
            "error"
        );
        establecerEstadoBotones({
            registrar: true,
            editar: false,
            eliminar: false
        });
        return;
    }

    /*
     * La confirmación evita eliminar una asignación por accidente.
     */
    if (!window.confirm(
        `¿Eliminar la asignación de "${asignacion.nombre}" ` +
        `al equipo "${equipo.Nombre}"?`
    )) {
        return;
    }

    establecerBotonOcupado(
        el.eliminar,
        true,
        "Eliminando...",
        "Eliminar"
    );

    try {
        await borrarAsignacion(asignacion.itemId);

        quitarAsignacionLocal(equipo.Nombre);

        asignacionGestionSeleccionada = null;

        if (el.trabajador) {
            el.trabajador.value = "";
            el.trabajador.disabled = false;
        }

        establecerEstadoBotones({
            registrar: true,
            editar: false,
            eliminar: false
        });

        mostrarMensaje(
            "Asignación eliminada correctamente.",
            "success"
        );
    } catch (error) {
        mostrarError(error, "No se pudo eliminar el trabajador");
    } finally {
        establecerBotonOcupado(
            el.eliminar,
            false,
            "Eliminando...",
            "Eliminar"
        );
    }
}

/* ========================= MONITOREO DE SERVICIOS ========================= */

function obtenerCampo(fields, nombres, defecto = "") {
    const entradas = Object.entries(fields || {});

    for (const nombre of nombres) {
        const encontrado = entradas.find(([clave]) =>
            normalizarClave(clave) === normalizarClave(nombre)
        );

        if (
            encontrado &&
            encontrado[1] !== null &&
            encontrado[1] !== undefined &&
            String(encontrado[1]).trim() !== ""
        ) {
            return encontrado[1];
        }
    }

    return defecto;
}

function formatearFecha(valor) {
    if (!valor) return "--";

    const fecha = new Date(valor);

    if (Number.isNaN(fecha.getTime())) {
        return String(valor);
    }

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

async function obtenerServicios() {
    const sitioActual = await obtenerSitio();

    const items = await obtenerTodosLosElementos(
        `https://graph.microsoft.com/v1.0/sites/${sitioActual.id}/lists/${encodeURIComponent(CONFIG.listaServicios)}/items?expand=fields&$top=200`
    );

    return items;
}

function renderizarServicios(items) {
    if (!el.serviciosGrid) return;

    if (!items.length) {
        el.serviciosGrid.innerHTML =
            '<div class="service-card">No hay servicios registrados.</div>';

        if (el.estadoGeneral) {
            el.estadoGeneral.textContent = "● SIN DATOS";
            el.estadoGeneral.className = "section-status warning";
        }

        return;
    }

    let online = 0;

    el.serviciosGrid.innerHTML = items.map(item => {
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

        const descripcion = obtenerCampo(
            fields,
            [
                "Descripcion",
                "Description",
                "Detalle",
                "Mensaje",
                "Observacion",
                "Comentario"
            ],
            "Sin descripción"
        );

        const latencia = obtenerCampo(
            fields,
            ["Latencia", "Ping", "Tiempo"],
            "--"
        );

        const revision = formatearFecha(obtenerCampo(
            fields,
            [
                "UltimaRevision",
                "Ultima revision",
                "Ultima_revision",
                "Fecha",
                "FechaRevision"
            ],
            ""
        ));

        const estadoNormalizado = normalizarClave(estado);
        let clase = "neutral";

        if (
            ["operativo", "online", "conectado"].some(valor =>
                estadoNormalizado.includes(valor)
            )
        ) {
            clase = "online";
            online++;
        } else if (
            ["mantenimiento", "advertencia", "warning", "lento"].some(valor =>
                estadoNormalizado.includes(valor)
            )
        ) {
            clase = "warning";
        } else if (
            ["caido", "offline", "desconectado", "error"].some(valor =>
                estadoNormalizado.includes(valor)
            )
        ) {
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
                        <span>Última revisión</span>
                        <strong>${escaparHTML(revision)}</strong>
                    </div>
                </div>
            </article>
        `;
    }).join("");

    if (el.estadoGeneral) {
        const clase = online === items.length
            ? "online"
            : online > 0 ? "warning" : "offline";

        el.estadoGeneral.textContent = online === items.length
            ? "● OPERATIVO"
            : online > 0 ? "● REVISAR" : "● NO DISPONIBLE";

        el.estadoGeneral.className = `section-status ${clase}`;
    }
}

async function cargarServicios() {
    try {
        renderizarServicios(await obtenerServicios());
    } catch (error) {
        console.error("[ALFERZA TI] Error cargando servicios:", error);

        if (el.estadoGeneral) {
            el.estadoGeneral.textContent = "● ERROR";
            el.estadoGeneral.className = "section-status offline";
        }
    }
}

/* ========================= TABLA DEL INVENTARIO ========================= */

function actualizarResumenEquipos() {
    const total = equiposData.length;

    const conectados = equiposData.filter(equipo =>
        normalizarClave(equipo.Estado) === "conectado"
    ).length;

    const desconectados = equiposData.filter(equipo =>
        normalizarClave(equipo.Estado) === "desconectado"
    ).length;

    if (el.total) el.total.textContent = total;
    if (el.conectados) el.conectados.textContent = conectados;
    if (el.desconectados) el.desconectados.textContent = desconectados;
}

function actualizarFiltroAreas() {
    if (!el.filtroArea) return;

    const seleccionAnterior = el.filtroArea.value;

    const areas = [...new Set(
        equiposData.map(equipo => equipo.Area).filter(Boolean)
    )].sort((a, b) =>
        String(a).localeCompare(String(b), "es")
    );

    el.filtroArea.innerHTML =
        '<option value="todos">Todas las áreas</option>';

    areas.forEach(area => {
        const opcion = document.createElement("option");
        opcion.value = area;
        opcion.textContent = area;
        el.filtroArea.appendChild(opcion);
    });

    if ([...el.filtroArea.options].some(
        opcion => opcion.value === seleccionAnterior
    )) {
        el.filtroArea.value = seleccionAnterior;
    }
}

function aplicarFiltrosEquipos(conservarPagina = false) {
    const texto = normalizar(el.buscar?.value);
    const estado = el.filtroEstado?.value || "todos";
    const area = el.filtroArea?.value || "todos";

    equiposFiltrados = equiposData.filter(equipo => {
        const coincideTexto = !texto || [
            equipo.Nombre,
            equipo.IP,
            equipo.DNSHostName
        ].some(valor => normalizar(valor).includes(texto));

        const coincideEstado = estado === "todos" ||
            String(equipo.Estado || "") === estado;

        const coincideArea = area === "todos" ||
            String(equipo.Area || "") === area;

        return coincideTexto && coincideEstado && coincideArea;
    });

    if (!conservarPagina) paginaActual = 1;

    renderizarTablaEquipos();
}

function claseEstado(estado) {
    const valor = normalizarClave(estado);

    if (valor === "conectado") return "online";
    if (valor === "desconectado") return "offline";

    return "neutral";
}

function renderizarTablaEquipos() {
    if (!el.tabla) return;

    const total = equiposFiltrados.length;
    const totalPaginas = Math.max(
        1,
        Math.ceil(total / CONFIG.equiposPorPagina)
    );

    paginaActual = Math.min(paginaActual, totalPaginas);

    const inicio = (paginaActual - 1) * CONFIG.equiposPorPagina;
    const pagina = equiposFiltrados.slice(
        inicio,
        inicio + CONFIG.equiposPorPagina
    );

    if (el.resultados) {
        el.resultados.textContent = total
            ? `Equipos encontrados: ${total}`
            : "0 equipos encontrados";
    }

    if (!total) {
        el.tabla.innerHTML = "";

        if (el.vacio) el.vacio.style.display = "flex";

        renderizarPaginacion(0);
        return;
    }

    if (el.vacio) el.vacio.style.display = "none";

    el.tabla.innerHTML = pagina.map(equipo => {
        const estado = equipo.Estado || "No verificable";
        const trabajador = obtenerAsignacion(equipo.Nombre);

        const trabajadorHTML = trabajador
            ? `<span class="trabajador-name" title="${escaparHTML(trabajador.nombre)}">${escaparHTML(trabajador.nombre)}</span>`
            : `<span class="trabajador-sin-asignar">Sin asignar</span>`;

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
                    <span class="equipo-status ${claseEstado(estado)}">
                        <span class="equipo-status-dot"></span>
                        ${escaparHTML(estado)}
                    </span>
                </td>
                <td class="trabajador-cell">${trabajadorHTML}</td>
                <td><span class="equipo-ip">${escaparHTML(equipo.IP || "--")}</span></td>
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
                    <span class="equipo-date">
                        ${escaparHTML(equipo.UltimoRegistroAD || "--")}
                    </span>
                </td>
            </tr>
        `;
    }).join("");

    renderizarPaginacion(totalPaginas);
}

function generarPaginas(actual, total) {
    if (total <= 7) {
        return Array.from({ length: total }, (_, i) => i + 1);
    }

    const paginas = [1];

    if (actual > 4) paginas.push("...");

    for (
        let pagina = Math.max(2, actual - 1);
        pagina <= Math.min(total - 1, actual + 1);
        pagina++
    ) {
        paginas.push(pagina);
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

    function agregarBoton(texto, pagina, deshabilitado = false, activo = false) {
        const boton = document.createElement("button");

        boton.type = "button";
        boton.className = "pagination-btn";
        boton.textContent = texto;
        boton.disabled = deshabilitado;

        if (activo) boton.classList.add("active");

        boton.addEventListener("click", () => {
            paginaActual = pagina;
            renderizarTablaEquipos();
        });

        contenedor.appendChild(boton);
    }

    agregarBoton("‹", paginaActual - 1, paginaActual === 1);

    generarPaginas(paginaActual, totalPaginas).forEach(pagina => {
        if (pagina === "...") {
            const span = document.createElement("span");
            span.className = "pagination-ellipsis";
            span.textContent = "...";
            contenedor.appendChild(span);
        } else {
            agregarBoton(
                String(pagina),
                pagina,
                false,
                pagina === paginaActual
            );
        }
    });

    agregarBoton(
        "›",
        paginaActual + 1,
        paginaActual === totalPaginas
    );

    const tabla = document.querySelector(".equipos-table-container");

    if (tabla?.parentNode) {
        tabla.parentNode.insertBefore(
            contenedor,
            tabla.nextSibling
        );
    }
}

function aplicarDatosEquipos(equipos) {
    equiposData = Array.isArray(equipos) ? equipos : [];

    actualizarResumenEquipos();
    actualizarFiltroAreas();
    aplicarFiltrosEquipos(true);
}

/* ========================= PERMISOS ========================= */

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
        } catch {
            return false;
        }
    }

    return usuarioPuedeVerInventario();
}

function configurarAccesoInventario() {
    const inventario = $("inventarioTISection");

    if (!inventario) return;

    inventario.style.display = "none";

    const actualizar = () => {
        inventario.style.display = usuarioPuedeVerInventario() ? "" : "none";
    };

    window.addEventListener("alferza:access-granted", actualizar);

    if (window.alferzaAccessReady?.then) {
        window.alferzaAccessReady.then(actualizar).catch(() => {
            inventario.style.display = "none";
        });
    } else {
        actualizar();
    }
}

/* ========================= CARGA DEL INVENTARIO ========================= */

async function cargarEquipos() {
    if (cargandoEquipos) return;

    cargandoEquipos = true;

    try {
        if (el.estadoInventario) {
            el.estadoInventario.textContent = "● ACTUALIZANDO";
            el.estadoInventario.className = "section-status warning";
        }

        const equipos = await obtenerEquiposSharePoint();

        if (equipos) {
            aplicarDatosEquipos(equipos);

            cacheGuardar("equipos", {
                etag: etagEquipos,
                datos: equipos
            });
        }

        if (el.ultimaRevision) {
            el.ultimaRevision.textContent = fechaActual();
        }

        if (el.estadoInventario) {
            el.estadoInventario.textContent = "● ACTUALIZADO";
            el.estadoInventario.className = "section-status online";
        }
    } catch (error) {
        console.error("[ALFERZA TI] Error cargando inventario:", error);

        if (el.estadoInventario) {
            el.estadoInventario.textContent = equiposData.length
                ? "● SIN ACTUALIZAR"
                : "● ERROR";

            el.estadoInventario.className = "section-status warning";
        }

        if (!equiposData.length && el.resultados) {
            el.resultados.textContent =
                "No se pudo cargar el inventario. Revisa la sesión y los permisos.";
        }
    } finally {
        cargandoEquipos = false;
    }
}

/* ========================= EVENTOS ========================= */

function configurarEventos() {
    el.abrirGestion?.addEventListener(
        "click",
        abrirGestionTrabajadores
    );

    el.buscarIP?.addEventListener(
        "click",
        buscarEquipoPorIP
    );

    el.ip?.addEventListener("keydown", event => {
        if (event.key === "Enter") {
            event.preventDefault();
            buscarEquipoPorIP();
        }

        if (event.key === "Escape") {
            cerrarGestionTrabajadores();
        }
    });

    el.trabajador?.addEventListener("keydown", event => {
        if (event.key === "Escape") {
            cerrarGestionTrabajadores();
        }

        if (event.key === "Enter") {
            event.preventDefault();

            if (!el.registrar?.disabled) {
                registrarTrabajadorPorIP();
            } else if (!el.editar?.disabled) {
                editarTrabajadorPorIP();
            }
        }
    });

    el.registrar?.addEventListener(
        "click",
        registrarTrabajadorPorIP
    );

    el.editar?.addEventListener(
        "click",
        editarTrabajadorPorIP
    );

    el.eliminar?.addEventListener(
        "click",
        eliminarTrabajadorPorIP
    );

    el.cerrar?.addEventListener(
        "click",
        cerrarGestionTrabajadores
    );

    el.cancelar?.addEventListener(
        "click",
        cerrarGestionTrabajadores
    );

    el.modal?.addEventListener("click", event => {
        if (event.target === el.modal) {
            cerrarGestionTrabajadores();
        }
    });

    el.buscar?.addEventListener("input", () => {
        clearTimeout(temporizadorBusqueda);

        temporizadorBusqueda = setTimeout(
            () => aplicarFiltrosEquipos(),
            200
        );
    });

    el.filtroEstado?.addEventListener(
        "change",
        () => aplicarFiltrosEquipos()
    );

    el.filtroArea?.addEventListener(
        "change",
        () => aplicarFiltrosEquipos()
    );
}

/* ========================= SESIÓN Y ACTUALIZACIÓN ========================= */

function mostrarBotonLogin() {
    if (el.loading) el.loading.style.display = "flex";

    if (el.loadingMessage) {
        el.loadingMessage.textContent =
            "Inicia sesión con Microsoft para consultar los datos.";
    }

    if ($("btnLoginMs") || !el.loading) return;

    const boton = document.createElement("button");

    boton.id = "btnLoginMs";
    boton.type = "button";
    boton.className = "loading-btn";
    boton.textContent = "Iniciar sesión con Microsoft";

    boton.addEventListener("click", async () => {
        boton.disabled = true;

        try {
            const token = await iniciarSesion(true);

            if (token) {
                boton.remove();
                ocultarCarga();
                await iniciarCargas();
            }
        } catch (error) {
            console.error("[ALFERZA TI] Error de inicio de sesión:", error);

            if (el.loadingMessage) {
                el.loadingMessage.textContent =
                    error.message || "No se pudo iniciar sesión.";
            }
        } finally {
            boton.disabled = false;
        }
    });

    const caja = el.loading.querySelector(".loading-content") ||
        el.loading.querySelector(".loading-card");

    caja?.appendChild(boton);
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
                console.error(
                    "[ALFERZA TI] Error cargando trabajadores:",
                    error
                );
            }
        }
    }

    ocultarCarga();
}

async function cargarDashboard() {
    try {
        if (!msalApp) {
            if (el.loadingMessage) {
                el.loadingMessage.textContent =
                    "No se pudo cargar la biblioteca de Microsoft.";
            }

            mostrarBotonLogin();
            return;
        }

        const cache = cacheLeer("equipos");

        if (cache && Array.isArray(cache.datos)) {
            etagEquipos = cache.etag || null;
            aplicarDatosEquipos(cache.datos);
        }

        const token = await iniciarSesion(false);

        if (!token) {
            mostrarBotonLogin();
            return;
        }

        ocultarCarga();
        await iniciarCargas();
    } catch (error) {
        console.error("[ALFERZA TI] Error iniciando dashboard:", error);
        mostrarBotonLogin();
    }
}

async function actualizarDashboard() {
    if (document.hidden) return;

    try {
        const token = await obtenerTokenSilencioso();

        if (!token) return;

        ultimaActualizacion = Date.now();

        await cargarServicios();

        if (await verificarPermisosInventario()) {
            await cargarEquipos();
        }
    } catch (error) {
        console.error("[ALFERZA TI] Error actualizando dashboard:", error);
    }
}

/* ========================= INICIALIZACIÓN ========================= */

function iniciarInfraestructura() {
    configurarAccesoInventario();
    configurarEventos();
    cargarDashboard();

    setInterval(
        actualizarDashboard,
        CONFIG.intervalo
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
        iniciarInfraestructura,
        { once: true }
    );
} else {
    iniciarInfraestructura();
}
