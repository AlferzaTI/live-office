/* =========================================
   INFRAESTRUCTURA - MONITOREO TI
========================================= */


/* =========================================
   CONFIGURACIÓN MICROSOFT
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

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


/* =========================================
   INSTANCIA MSAL
========================================= */

if (
    typeof msal === "undefined"
) {

    console.error(
        "❌ MSAL no está disponible."
    );

} else {

    console.log(
        "✅ MSAL disponible en Infraestructura."
    );

}


const msalInstanceInfraestructura =
    new msal.PublicClientApplication(
        msalConfigInfraestructura
    );


/* =========================================
   CONFIGURACIÓN SHAREPOINT
========================================= */

const SHAREPOINT_HOSTNAME =
    "alferzaholding-my.sharepoint.com";

const SHAREPOINT_SITE =
    "/personal/soporte1_alferza_pe";

const SHAREPOINT_LIST =
    "MonitoreoTI";


/* =========================================
   ELEMENTOS
========================================= */

const loadingOverlay =
    document.getElementById(
        "loadingOverlay"
    );


const loadingMessage =
    document.getElementById(
        "loadingMessage"
    );


/* =========================================
   VARIABLES DE EQUIPOS
========================================= */

let equiposData = [];


/* =========================================
   PAGINACIÓN
========================================= */

const EQUIPOS_POR_PAGINA =
    40;

let paginaEquiposActual =
    1;

let equiposFiltradosActuales =
    [];


/* =========================================
   MENSAJE DE CARGA
========================================= */

function cambiarMensaje(
    mensaje
) {

    if (
        loadingMessage
    ) {

        loadingMessage.textContent =
            mensaje;

    }

}


/* =========================================
   MOSTRAR CARGA
========================================= */

function mostrarCarga(
    mensaje
) {

    cambiarMensaje(
        mensaje
    );


    if (
        loadingOverlay
    ) {

        loadingOverlay.classList.remove(
            "hidden"
        );

        loadingOverlay.classList.remove(
            "oculto"
        );

    }

}


/* =========================================
   OCULTAR CARGA
========================================= */

function ocultarCarga() {

    if (
        loadingOverlay
    ) {

        loadingOverlay.classList.add(
            "hidden"
        );

    }

}


/* =========================================
   OBTENER TOKEN
========================================= */

async function obtenerTokenInfraestructura() {

    console.log(
        "🔐 Buscando sesión Microsoft..."
    );


    const cuentas =
        msalInstanceInfraestructura
            .getAllAccounts();


    console.log(
        "Cuentas encontradas:",
        cuentas.length
    );


    if (
        !cuentas.length
    ) {

        throw new Error(
            "No se encontró una cuenta Microsoft activa."
        );

    }


    const cuenta =
        cuentas[0];


    msalInstanceInfraestructura
        .setActiveAccount(
            cuenta
        );


    console.log(
        "👤 Cuenta activa:",
        cuenta.username
    );


    try {

        cambiarMensaje(
            "Verificando permisos..."
        );


        const respuesta =
            await msalInstanceInfraestructura
                .acquireTokenSilent({

                    scopes: [

                        "User.Read",

                        "Sites.Read.All"

                    ],

                    account:
                        cuenta

                });


        console.log(
            "✅ Token obtenido correctamente."
        );


        return respuesta.accessToken;

    }

    catch (
        error
    ) {

        console.error(
            "❌ Error obteniendo token silenciosamente:",
            error
        );

        throw error;

    }

}


/* =========================================
   OBTENER SITIO SHAREPOINT
========================================= */

async function obtenerSitioSharePoint(
    TOKEN
) {

    console.log(
        "🔎 Buscando sitio SharePoint..."
    );


    const respuesta =
        await fetch(

            `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOSTNAME}:${SHAREPOINT_SITE}`,

            {

                method:
                    "GET",

                headers: {

                    Authorization:
                        `Bearer ${TOKEN}`

                }

            }

        );


    if (
        !respuesta.ok
    ) {

        const error =
            await respuesta.text();


        console.error(
            "❌ Error obteniendo sitio SharePoint:",
            error
        );


        throw new Error(
            `SharePoint Site: HTTP ${respuesta.status}`
        );

    }


    const sitio =
        await respuesta.json();


    console.log(
        "✅ SharePoint encontrado:",
        sitio
    );


    return sitio;

}


/* =========================================
   OBTENER LISTA MONITOREOTI
========================================= */

/* =========================================
   OBTENER MONITOREO TI
========================================= */

async function obtenerMonitoreoTI(TOKEN, sitio) {

    console.log("🔎 Consultando MonitoreoTI...");

    const respuesta = await fetch(

        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${SHAREPOINT_LIST}/items?expand=fields&$top=999`,

        {
            method: "GET",

            headers: {
                Authorization: `Bearer ${TOKEN}`,
                Accept: "application/json"
            }
        }

    );


    if (!respuesta.ok) {

        const error =
            await respuesta.text();

        console.error(
            "❌ Error obteniendo MonitoreoTI:",
            respuesta.status,
            error
        );

        throw new Error(
            `MonitoreoTI HTTP ${respuesta.status}`
        );

    }


    const data =
        await respuesta.json();


    console.log(
        "✅ MonitoreoTI:",
        data.value?.length || 0,
        "registros"
    );


    return data.value || [];

}


/* =========================================
   OBTENER EQUIPOS
========================================= */

async function obtenerEquipos(TOKEN, sitio) {

    console.log(
        "🖥️ Cargando equipos.json..."
    );


    const respuestaDrive =
        await fetch(

            `https://graph.microsoft.com/v1.0/sites/${sitio.id}/drive`,

            {
                method: "GET",

                headers: {
                    Authorization: `Bearer ${TOKEN}`,
                    Accept: "application/json"
                }
            }

        );


    if (!respuestaDrive.ok) {

        const error =
            await respuestaDrive.text();

        console.error(
            "❌ Error obteniendo Drive:",
            respuestaDrive.status,
            error
        );

        throw new Error(
            `Drive HTTP ${respuestaDrive.status}`
        );

    }


    const drive =
        await respuestaDrive.json();


    if (!drive.id) {

        throw new Error(
            "El Drive de SharePoint no tiene ID."
        );

    }


    const rutaArchivo =
        "Procedimientos T.I/equipos.json";


    const urlArchivo =
        `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${encodeURI(rutaArchivo)}:/content`;


    console.log(
        "📂 Descargando:",
        rutaArchivo
    );


    const respuestaArchivo =
        await fetch(

            urlArchivo,

            {
                method: "GET",

                headers: {
                    Authorization: `Bearer ${TOKEN}`,
                    Accept: "application/json"
                }
            }

        );


    if (!respuestaArchivo.ok) {

        const error =
            await respuestaArchivo.text();

        console.error(
            "❌ Error obteniendo equipos.json:",
            respuestaArchivo.status,
            error
        );

        throw new Error(
            `equipos.json HTTP ${respuestaArchivo.status}`
        );

    }


    const datos =
        await respuestaArchivo.json();


    if (!Array.isArray(datos)) {

        throw new Error(
            "equipos.json no contiene un array válido."
        );

    }


    console.log(
        "✅ Equipos encontrados:",
        datos.length
    );


    return datos;

}


/* =========================================
   CARGAR EQUIPOS
========================================= */

async function cargarEquipos(
    TOKEN = null,
    sitio = null
) {

    try {

        console.log(
            "===================================="
        );

        console.log(
            "🖥️ INICIANDO INVENTARIO"
        );

        console.log(
            "===================================="
        );


        /*
         * Si no recibimos token/sitio,
         * los obtenemos.
         */

        if (!TOKEN) {

            TOKEN =
                await obtenerTokenInfraestructura();

        }


        if (!sitio) {

            sitio =
                await obtenerSitioSharePoint(
                    TOKEN
                );

        }


        const equipos =
            await obtenerEquipos(
                TOKEN,
                sitio
            );


        equiposData =
            equipos;


        actualizarResumenEquipos(
            equiposData
        );


        cargarFiltroAreas(
            equiposData
        );


        paginaEquiposActual =
            1;


        renderizarEquipos(
            equiposData
        );


        configurarFiltrosEquipos();


        const estado =
            document.getElementById(
                "equiposEstado"
            );


        if (estado) {

            estado.textContent =
                "● ACTUALIZADO";

            estado.className =
                "section-status online";

        }


        console.log(
            "✅ Inventario cargado:",
            equiposData.length,
            "equipos"
        );


        return equiposData;

    }

    catch (error) {

        console.error(
            "❌ Error cargando inventario:",
            error
        );


        const tbody =
            document.getElementById(
                "equiposTableBody"
            );


        if (tbody) {

            tbody.innerHTML = `

                <tr>

                    <td
                        colspan="7"
                        class="equipos-loading"
                    >

                        ❌ No se pudo cargar
                        <strong>el inventario desde SharePoint</strong>.

                    </td>

                </tr>

            `;

        }


        const estado =
            document.getElementById(
                "equiposEstado"
            );


        if (estado) {

            estado.textContent =
                "● ERROR";

            estado.className =
                "section-status offline";

        }


        /*
         * Importante:
         * no propagamos el error.
         *
         * Así una falla del inventario
         * no bloquea todo el dashboard.
         */

        return [];

    }

}


/* =========================================
   CARGAR SERVICIOS
========================================= */

async function cargarServicios(
    TOKEN,
    sitio
) {

    try {

        console.log(
            "☁️ Cargando servicios..."
        );


        const datos =
            await obtenerMonitoreoTI(
                TOKEN,
                sitio
            );


        renderizarServicios(
            datos
        );


        actualizarResumen(
            datos
        );


        actualizarEstadoGeneral(
            datos
        );


        actualizarHora(
            datos
        );


        console.log(
            "✅ Servicios cargados:",
            datos.length
        );


        return datos;

    }

    catch (error) {

        console.error(
            "❌ Error cargando servicios:",
            error
        );


        const grid =
            document.getElementById(
                "serviciosGrid"
            );


        if (grid) {

            grid.innerHTML = `

                <div class="empty-state">

                    ❌ No se pudieron cargar
                    los servicios.

                </div>

            `;

        }


        const estado =
            document.getElementById(
                "estadoGeneral"
            );


        if (estado) {

            estado.textContent =
                "● ERROR";

            estado.className =
                "section-status offline";

        }


        return [];

    }

}


/* =========================================
   CARGAR TODO EL DASHBOARD
========================================= */

async function cargarMonitoreo() {

    console.log(
        "===================================="
    );

    console.log(
        "🚀 INICIANDO CENTRO DE MONITOREO TI"
    );

    console.log(
        "===================================="
    );


    mostrarCarga(
        "Conectando con Microsoft..."
    );


    const inicio =
        performance.now();


    try {

        /* =====================================
           1. TOKEN — SOLO UNA VEZ
        ===================================== */

        const TOKEN =
            await obtenerTokenInfraestructura();


        console.log(
            "🔐 Token listo."
        );


        /* =====================================
           2. SHAREPOINT — SOLO UNA VEZ
        ===================================== */

        cambiarMensaje(
            "Conectando con SharePoint..."
        );


        const sitio =
            await obtenerSitioSharePoint(
                TOKEN
            );


        console.log(
            "📍 Sitio listo:",
            sitio.id
        );


        /* =====================================
           3. SERVICIOS + EQUIPOS EN PARALELO
        ===================================== */

        cambiarMensaje(
            "Cargando monitoreo..."
        );


        const resultado =
            await Promise.allSettled([

                cargarServicios(
                    TOKEN,
                    sitio
                ),

                cargarEquipos(
                    TOKEN,
                    sitio
                )

            ]);


        console.log(
            "📊 Resultado de cargas:",
            resultado
        );


        /* =====================================
           4. TIEMPO TOTAL
        ===================================== */

        const tiempo =
            (
                performance.now() -
                inicio
            ) / 1000;


        console.log(
            `⚡ Dashboard cargado en ${tiempo.toFixed(2)} segundos`
        );


    }

    catch (error) {

        console.error(
            "❌ Error general del dashboard:",
            error
        );


        cambiarMensaje(
            "No se pudo conectar con Microsoft."
        );

    }

    finally {

        /*
         * SIEMPRE ocultamos el overlay.
         *
         * Aunque Graph o SharePoint
         * tenga un problema.
         */

        setTimeout(
            ocultarCarga,
            300
        );

    }

}


/* =========================================
   ACTUALIZACIÓN AUTOMÁTICA
========================================= */

let actualizacionEnCurso =
    false;


async function actualizarDashboard() {

    /*
     * Evita que una actualización nueva
     * empiece mientras la anterior sigue
     * ejecutándose.
     */

    if (actualizacionEnCurso) {

        console.log(
            "⏳ Ya existe una actualización en curso."
        );

        return;

    }


    actualizacionEnCurso =
        true;


    try {

        console.log(
            "🔄 Actualización automática..."
        );


        const TOKEN =
            await obtenerTokenInfraestructura();


        const sitio =
            await obtenerSitioSharePoint(
                TOKEN
            );


        await Promise.allSettled([

            cargarServicios(
                TOKEN,
                sitio
            ),

            cargarEquipos(
                TOKEN,
                sitio
            )

        ]);


        console.log(
            "✅ Actualización terminada."
        );

    }

    catch (error) {

        console.error(
            "❌ Error en actualización automática:",
            error
        );

    }

    finally {

        actualizacionEnCurso =
            false;

    }

}


/* =========================================
   INICIO
========================================= */

document.addEventListener(
    "DOMContentLoaded",
    function () {

        console.log(
            "🚀 DOM listo."
        );


        cargarMonitoreo();

    }
);


/* =========================================
   ACTUALIZACIÓN CADA 2 MINUTOS
========================================= */

setInterval(
    actualizarDashboard,
    120000
);

/* =========================================
   OBTENER CAMPOS
========================================= */

function obtenerCampos(
    item
) {

    return (
        item.fields || {}
    );

}


/* =========================================
   NORMALIZAR ESTADO
========================================= */

function normalizarEstado(
    estado
) {

    if (
        !estado
    ) {

        return "No configurado";

    }


    return String(
        estado
    ).trim();

}


/* =========================================
   CLASIFICAR ESTADO
========================================= */

function clasificarEstado(
    estado
) {

    const valor =
        normalizarEstado(
            estado
        ).toLowerCase();


    if (
        valor.includes("operativo") ||
        valor.includes("vigente") ||
        valor.includes("activo") ||
        valor.includes("responde")
    ) {

        return "online";

    }


    if (
        valor.includes("advertencia") ||
        valor.includes("warning") ||
        valor.includes("pendiente")
    ) {

        return "warning";

    }


    if (
        valor.includes("incidencia") ||
        valor.includes("error") ||
        valor.includes("caído") ||
        valor.includes("caido")
    ) {

        return "offline";

    }


    return "neutral";

}


/* =========================================
   ICONO DEL SERVICIO
========================================= */

function obtenerIconoServicio(
    servicio
) {

    const nombre =
        String(
            servicio
        ).toLowerCase();


    if (
        nombre.includes("internet")
    ) {

        return "🌐";

    }


    if (
        nombre.includes("microsoft")
    ) {

        return "☁️";

    }


    if (
        nombre.includes("certificado")
    ) {

        return "🔒";

    }


    if (
        nombre.includes("servidor")
    ) {

        return "🖥️";

    }


    if (
        nombre.includes("dns")
    ) {

        return "🔎";

    }


    return "⚙️";

}


/* =========================================
   FORMATEAR ÚLTIMA REVISIÓN
========================================= */

function formatearUltimaRevision(
    fecha
) {

    if (
        !fecha
    ) {

        return "Sin registro";

    }


    const fechaObjeto =
        new Date(
            fecha
        );


    if (
        Number.isNaN(
            fechaObjeto.getTime()
        )
    ) {

        return String(
            fecha
        );

    }


    return fechaObjeto.toLocaleString(
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
                "2-digit",

            second:
                "2-digit"

        }
    );

}


/* =========================================
   ACTUALIZAR ÚLTIMA ACTUALIZACIÓN
========================================= */

function actualizarHora(
    datos
) {

    const elemento =
        document.getElementById(
            "lastUpdate"
        );


    if (
        !elemento
    ) return;


    if (
        !datos ||
        datos.length === 0
    ) {

        elemento.textContent =
            "Sin registro";

        return;

    }


    const fechas =
        datos
            .map(
                item => {

                    const campos =
                        obtenerCampos(
                            item
                        );

                    return campos.Ultimarevision;

                }
            )
            .filter(
                fecha =>
                    fecha
            );


    if (
        fechas.length === 0
    ) {

        elemento.textContent =
            "Sin registro";

        return;

    }


    const fechasValidas =
        fechas
            .map(
                fecha =>
                    new Date(
                        fecha
                    )
            )
            .filter(
                fecha =>
                    !Number.isNaN(
                        fecha.getTime()
                    )
            );


    if (
        fechasValidas.length === 0
    ) {

        elemento.textContent =
            "Fecha inválida";

        return;

    }


    const ultimaFecha =
        new Date(
            Math.max(
                ...fechasValidas.map(
                    fecha =>
                        fecha.getTime()
                )
            )
        );


    elemento.textContent =
        formatearUltimaRevision(
            ultimaFecha
        );


    console.log(
        "🕐 Última revisión general:",
        elemento.textContent
    );

}


/* =========================================
   RENDERIZAR SERVICIOS
========================================= */

function renderizarServicios(
    datos
) {

    const grid =
        document.getElementById(
            "serviciosGrid"
        );


    const estadoGeneral =
        document.getElementById(
            "estadoGeneral"
        );


    if (
        !grid
    ) return;


    grid.innerHTML =
        "";


    if (
        !datos ||
        datos.length === 0
    ) {

        grid.innerHTML = `
            <div class="empty-state">
                No hay servicios registrados.
            </div>
        `;


        if (
            estadoGeneral
        ) {

            estadoGeneral.textContent =
                "● NO CONFIGURADO";

            estadoGeneral.className =
                "section-status";

        }


        return;

    }


    const estados = [];


    datos.forEach(
        item => {

            const campos =
                obtenerCampos(
                    item
                );


            const servicio =
                campos.Servicio ||
                "Servicio";


            const estado =
                campos.Estado ||
                "No configurado";


            const detalle =
                campos.Detalle ||
                "Sin información disponible";


            const latencia =
                campos.Latencia;


            const ultimaRevision =
                campos.Ultimarevision;


            const estadoNormalizado =
                normalizarEstado(
                    estado
                );


            estados.push(
                estadoNormalizado
            );


            const tipoEstado =
                clasificarEstado(
                    estado
                );


            const icono =
                obtenerIconoServicio(
                    servicio
                );


            const latenciaTexto =
                latencia !== undefined &&
                latencia !== null &&
                latencia !== ""
                    ? `${latencia} ms`
                    : "—";


            const revisionTexto =
                formatearUltimaRevision(
                    ultimaRevision
                );


            const card =
                document.createElement(
                    "div"
                );


            card.className =
                `service-card ${tipoEstado}`;


            /*
             * ESTRUCTURA HORIZONTAL:
             *
             * ICONO | INFORMACIÓN
             *
             * Esto permite que la tarjeta sea
             * mucho más compacta.
             */

            card.innerHTML = `

                <div class="service-main">

                    <div class="service-icon">
                        ${icono}
                    </div>

                    <div class="service-info">

                        <div class="service-heading">

                            <h3>
                                ${escaparHTML(servicio)}
                            </h3>

                            <div class="service-status ${tipoEstado}">

                                <span class="status-dot"></span>

                                ${escaparHTML(estado)}

                            </div>

                        </div>

                        <p>
                            ${escaparHTML(detalle)}
                        </p>

                    </div>

                </div>


                <div class="service-meta">

                    <div class="service-latency">

                        <span>
                            Latencia
                        </span>

                        <strong>
                            ${escaparHTML(latenciaTexto)}
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
                                ${escaparHTML(revisionTexto)}
                            </strong>

                        </div>

                    </div>

                </div>

            `;


            grid.appendChild(
                card
            );

        }
    );


    /* =====================================
       ESTADO GENERAL
    ===================================== */

    if (
        estadoGeneral
    ) {

        const estadosNormalizados =
            estados.map(
                estado =>
                    String(
                        estado
                    )
                    .trim()
                    .toLowerCase()
            );


        const hayIncidencia =
            estadosNormalizados.some(
                estado =>
                    [
                        "incidencia",
                        "error",
                        "caido",
                        "caído"
                    ].includes(
                        estado
                    )
            );


        const hayAdvertencia =
            estadosNormalizados.some(
                estado =>
                    [
                        "advertencia",
                        "warning",
                        "pendiente"
                    ].includes(
                        estado
                    )
            );


        const todosNoConfigurados =
            estadosNormalizados.length > 0 &&
            estadosNormalizados.every(
                estado =>
                    estado === "no configurado" ||
                    estado === ""
            );


        const todosOperativos =
            estadosNormalizados.length > 0 &&
            estadosNormalizados.every(
                estado =>
                    [
                        "operativo",
                        "vigente",
                        "activo",
                        "responde"
                    ].includes(
                        estado
                    )
            );


        if (
            hayIncidencia
        ) {

            estadoGeneral.textContent =
                "● INCIDENCIA";

            estadoGeneral.className =
                "section-status offline";

        }

        else if (
            hayAdvertencia
        ) {

            estadoGeneral.textContent =
                "● ADVERTENCIA";

            estadoGeneral.className =
                "section-status warning";

        }

        else if (
            todosOperativos
        ) {

            estadoGeneral.textContent =
                "● OPERATIVO";

            estadoGeneral.className =
                "section-status online";

        }

        else if (
            todosNoConfigurados
        ) {

            estadoGeneral.textContent =
                "● NO CONFIGURADO";

            estadoGeneral.className =
                "section-status";

        }

        else {

            estadoGeneral.textContent =
                "● NO CONFIGURADO";

            estadoGeneral.className =
                "section-status";

        }

    }

}


/* =========================================
   ACTUALIZAR RESUMEN
========================================= */

function actualizarResumen(
    datos
) {

    let activos =
        0;

    let advertencias =
        0;

    let incidencias =
        0;


    datos.forEach(
        item => {

            const fields =
                obtenerCampos(
                    item
                );


            const estado =
                clasificarEstado(
                    fields.Estado
                );


            if (
                estado === "online"
            ) {

                activos++;

            }

            else if (
                estado === "warning"
            ) {

                advertencias++;

            }

            else if (
                estado === "offline"
            ) {

                incidencias++;

            }

        }
    );


    const activosElemento =
        document.getElementById(
            "serviciosActivos"
        );


    const advertenciasElemento =
        document.getElementById(
            "advertencias"
        );


    const incidenciasElemento =
        document.getElementById(
            "incidencias"
        );


    if (
        activosElemento
    ) {

        activosElemento.textContent =
            activos;

    }


    if (
        advertenciasElemento
    ) {

        advertenciasElemento.textContent =
            advertencias;

    }


    if (
        incidenciasElemento
    ) {

        incidenciasElemento.textContent =
            incidencias;

    }

}


/* =========================================
   ACTUALIZAR ESTADO GENERAL
========================================= */

function actualizarEstadoGeneral(
    datos
) {

    const estadoGeneral =
        document.getElementById(
            "estadoMonitoreo"
        );


    if (
        !estadoGeneral
    ) return;


    if (
        !datos ||
        datos.length === 0
    ) {

        estadoGeneral.textContent =
            "No configurado";

        estadoGeneral.className =
            "status-badge neutral";

        return;

    }


    const estados =
        datos.map(
            item => {

                const campos =
                    obtenerCampos(
                        item
                    );

                return normalizarEstado(
                    campos.Estado
                )
                .toLowerCase();

            }
        );


    const tieneIncidencia =
        estados.some(
            estado =>
                [
                    "incidencia",
                    "error",
                    "caido",
                    "caído"
                ].includes(
                    estado
                )
        );


    const tieneAdvertencia =
        estados.some(
            estado =>
                [
                    "advertencia",
                    "warning",
                    "pendiente"
                ].includes(
                    estado
                )
        );


    const todosNoConfigurados =
        estados.every(
            estado =>
                estado === "no configurado" ||
                estado === ""
        );


    const todosOperativos =
        estados.every(
            estado =>
                [
                    "operativo",
                    "vigente",
                    "activo",
                    "responde"
                ].includes(
                    estado
                )
        );


    if (
        tieneIncidencia
    ) {

        estadoGeneral.textContent =
            "Incidencia";

        estadoGeneral.className =
            "status-badge offline";

    }

    else if (
        tieneAdvertencia
    ) {

        estadoGeneral.textContent =
            "Advertencia";

        estadoGeneral.className =
            "status-badge warning";

    }

    else if (
        todosNoConfigurados
    ) {

        estadoGeneral.textContent =
            "No configurado";

        estadoGeneral.className =
            "status-badge neutral";

    }

    else if (
        todosOperativos
    ) {

        estadoGeneral.textContent =
            "Operativo";

        estadoGeneral.className =
            "status-badge online";

    }

    else {

        estadoGeneral.textContent =
            "No configurado";

        estadoGeneral.className =
            "status-badge neutral";

    }

}


/* =========================================
   EQUIPOS
========================================= */


/* =========================================
   OBTENER EQUIPOS DESDE SHAREPOINT
========================================= */

async function obtenerEquipos() {

    try {

        console.log(
            "🔄 Cargando inventario desde SharePoint..."
        );


        const TOKEN =
            await obtenerTokenInfraestructura();


        if (
            !TOKEN
        ) {

            throw new Error(
                "No se obtuvo el token de Microsoft Graph."
            );

        }


        /* =====================================
           1. OBTENER SITIO
        ===================================== */

        const sitio =
            await obtenerSitioSharePoint(
                TOKEN
            );


        if (
            !sitio ||
            !sitio.id
        ) {

            throw new Error(
                "No se pudo obtener el sitio de SharePoint."
            );

        }


        /* =====================================
           2. OBTENER DRIVE
        ===================================== */

        const respuestaDrive =
            await fetch(

                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/drive`,

                {

                    method:
                        "GET",

                    headers: {

                        Authorization:
                            `Bearer ${TOKEN}`,

                        Accept:
                            "application/json"

                    }

                }

            );


        if (
            !respuestaDrive.ok
        ) {

            const errorDrive =
                await respuestaDrive.text();


            console.error(
                "❌ Error obteniendo Drive:",
                respuestaDrive.status,
                errorDrive
            );


            throw new Error(
                `No se pudo obtener el Drive (${respuestaDrive.status})`
            );

        }


        const drive =
            await respuestaDrive.json();


        if (
            !drive.id
        ) {

            throw new Error(
                "El Drive no tiene un ID válido."
            );

        }


        /* =====================================
           3. RUTA EQUIPOS.JSON
        ===================================== */

        const rutaArchivo =
            "Procedimientos T.I/equipos.json";


        const urlArchivo =
            `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${encodeURI(rutaArchivo)}:/content`;


        /* =====================================
           4. DESCARGAR EQUIPOS.JSON
        ===================================== */

        const respuestaArchivo =
            await fetch(

                urlArchivo,

                {

                    method:
                        "GET",

                    headers: {

                        Authorization:
                            `Bearer ${TOKEN}`,

                        Accept:
                            "application/json"

                    }

                }

            );


        if (
            !respuestaArchivo.ok
        ) {

            const errorArchivo =
                await respuestaArchivo.text();


            console.error(
                "❌ Error obteniendo equipos.json:",
                respuestaArchivo.status,
                errorArchivo
            );


            throw new Error(
                `No se pudo obtener equipos.json (${respuestaArchivo.status})`
            );

        }


        /* =====================================
           5. LEER JSON
        ===================================== */

        const datos =
            await respuestaArchivo.json();


        console.log(
            "📊 Equipos encontrados:",
            Array.isArray(datos)
                ? datos.length
                : 0
        );


        /* =====================================
           6. VALIDAR
        ===================================== */

        if (
            !Array.isArray(
                datos
            )
        ) {

            throw new Error(
                "El archivo equipos.json no tiene el formato esperado."
            );

        }


        if (
            datos.length === 0
        ) {

            console.warn(
                "⚠️ equipos.json está vacío."
            );


            return [];

        }


        console.log(
            "✅ Inventario cargado:",
            datos.length,
            "equipos"
        );


        return datos;

    }

    catch (
        error
    ) {

        console.error(
            "❌ Error cargando inventario desde SharePoint:",
            error
        );


        return [];

    }

}


/* =========================================
   FORMATEAR FECHA
========================================= */

function formatearFechaEquipo(
    fecha
) {

    if (
        !fecha
    ) {

        return "—";

    }


    const fechaObjeto =
        new Date(
            fecha
        );


    if (
        Number.isNaN(
            fechaObjeto.getTime()
        )
    ) {

        return String(
            fecha
        );

    }


    return fechaObjeto.toLocaleDateString(
        "es-PE",
        {

            day:
                "2-digit",

            month:
                "2-digit",

            year:
                "numeric"

        }
    );

}


/* =========================================
   ESCAPAR HTML
========================================= */

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


/* =========================================
   CLASIFICAR ESTADO EQUIPO
========================================= */

function clasificarEstadoEquipo(
    estado
) {

    const valor =
        String(
            estado || ""
        )
        .trim()
        .toLowerCase();


    if (
        valor === "conectado"
    ) {

        return "online";

    }


    if (
        valor === "desconectado"
    ) {

        return "offline";

    }


    return "neutral";

}


/* =========================================
   CREAR CONTENEDOR PAGINACIÓN
========================================= */

function obtenerContenedorPaginacion() {

    let contenedor =
        document.getElementById(
            "equiposPagination"
        );


    if (
        contenedor
    ) {

        return contenedor;

    }


    const tableContainer =
        document.querySelector(
            ".equipos-table-container"
        );


    if (
        !tableContainer
    ) {

        return null;

    }


    contenedor =
        document.createElement(
            "div"
        );


    contenedor.id =
        "equiposPagination";


    contenedor.className =
        "equipos-pagination";


    tableContainer.insertAdjacentElement(
        "afterend",
        contenedor
    );


    return contenedor;

}


/* =========================================
   RENDERIZAR PAGINACIÓN
========================================= */

function renderizarPaginacion(
    totalEquipos
) {

    const contenedor =
        obtenerContenedorPaginacion();


    if (
        !contenedor
    ) return;


    contenedor.innerHTML =
        "";


    const totalPaginas =
        Math.ceil(
            totalEquipos /
            EQUIPOS_POR_PAGINA
        );


    if (
        totalPaginas <= 1
    ) {

        contenedor.style.display =
            "none";

        return;

    }


    contenedor.style.display =
        "flex";


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


    botonAnterior.title =
        "Página anterior";


    botonAnterior.disabled =
        paginaEquiposActual === 1;


    botonAnterior.addEventListener(
        "click",
        function () {

            if (
                paginaEquiposActual > 1
            ) {

                paginaEquiposActual--;

                renderizarEquipos(
                    equiposFiltradosActuales,
                    false
                );

            }

        }
    );


    contenedor.appendChild(
        botonAnterior
    );


    /*
     * Generamos páginas visibles.
     *
     * Con muchos resultados no mostramos
     * 50 botones de golpe.
     */

    const paginas =
        obtenerPaginasVisibles(
            totalPaginas,
            paginaEquiposActual
        );


    paginas.forEach(
        pagina => {

            if (
                pagina === "..."
            ) {

                const separador =
                    document.createElement(
                        "span"
                    );


                separador.className =
                    "pagination-ellipsis";


                separador.textContent =
                    "…";


                contenedor.appendChild(
                    separador
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


            if (
                pagina === paginaEquiposActual
            ) {

                boton.classList.add(
                    "active"
                );

            }


            boton.textContent =
                pagina;


            boton.addEventListener(
                "click",
                function () {

                    paginaEquiposActual =
                        pagina;

                    renderizarEquipos(
                        equiposFiltradosActuales,
                        false
                    );

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


    botonSiguiente.title =
        "Página siguiente";


    botonSiguiente.disabled =
        paginaEquiposActual === totalPaginas;


    botonSiguiente.addEventListener(
        "click",
        function () {

            if (
                paginaEquiposActual <
                totalPaginas
            ) {

                paginaEquiposActual++;

                renderizarEquipos(
                    equiposFiltradosActuales,
                    false
                );

            }

        }
    );


    contenedor.appendChild(
        botonSiguiente
    );

}


/* =========================================
   PÁGINAS VISIBLES
========================================= */

function obtenerPaginasVisibles(
    totalPaginas,
    paginaActual
) {

    if (
        totalPaginas <= 7
    ) {

        return Array.from(
            {
                length:
                    totalPaginas
            },
            (
                _,
                indice
            ) =>
                indice + 1
        );

    }


    const paginas = [];


    paginas.push(
        1
    );


    if (
        paginaActual > 4
    ) {

        paginas.push(
            "..."
        );

    }


    const inicio =
        Math.max(
            2,
            paginaActual - 1
        );


    const fin =
        Math.min(
            totalPaginas - 1,
            paginaActual + 1
        );


    for (
        let pagina = inicio;
        pagina <= fin;
        pagina++
    ) {

        paginas.push(
            pagina
        );

    }


    if (
        paginaActual <
        totalPaginas - 3
    ) {

        paginas.push(
            "..."
        );

    }


    paginas.push(
        totalPaginas
    );


    return paginas;

}


/* =========================================
   ACTUALIZAR TEXTO DE RESULTADOS
========================================= */

function actualizarTextoResultados(
    totalEquipos,
    inicio,
    fin
) {

    const resultados =
        document.getElementById(
            "equiposResultados"
        );


    if (
        !resultados
    ) return;


    if (
        totalEquipos === 0
    ) {

        resultados.textContent =
            "0 equipos";

        return;

    }


    resultados.textContent =
        `Mostrando ${inicio}–${fin} de ${totalEquipos} equipos`;

}


/* =========================================
   RENDERIZAR TABLA EQUIPOS
========================================= */

function renderizarEquipos(
    equipos,
    actualizarPaginacion = true
) {

    const tbody =
        document.getElementById(
            "equiposTableBody"
        );


    const empty =
        document.getElementById(
            "equiposEmpty"
        );


    if (
        !tbody
    ) return;


    /*
     * Guardamos los resultados actuales
     * para poder cambiar de página.
     */

    equiposFiltradosActuales =
        equipos || [];


    tbody.innerHTML =
        "";


    /* =====================================
       SIN RESULTADOS
    ===================================== */

    if (
        !equipos ||
        equipos.length === 0
    ) {

        if (
            empty
        ) {

            empty.style.display =
                "block";

        }


        actualizarTextoResultados(
            0,
            0,
            0
        );


        const paginacion =
            obtenerContenedorPaginacion();


        if (
            paginacion
        ) {

            paginacion.style.display =
                "none";

        }


        return;

    }


    if (
        empty
    ) {

        empty.style.display =
            "none";

    }


    /* =====================================
       VALIDAR PÁGINA
    ===================================== */

    const totalPaginas =
        Math.ceil(
            equipos.length /
            EQUIPOS_POR_PAGINA
        );


    if (
        paginaEquiposActual >
        totalPaginas
    ) {

        paginaEquiposActual =
            totalPaginas;

    }


    if (
        paginaEquiposActual < 1
    ) {

        paginaEquiposActual =
            1;

    }


    /* =====================================
       CALCULAR RANGO
    ===================================== */

    const inicioIndice =
        (
            paginaEquiposActual -
            1
        ) *
        EQUIPOS_POR_PAGINA;


    const finIndice =
        Math.min(
            inicioIndice +
            EQUIPOS_POR_PAGINA,
            equipos.length
        );


    const equiposPagina =
        equipos.slice(
            inicioIndice,
            finIndice
        );


    const inicioVisible =
        inicioIndice + 1;


    const finVisible =
        finIndice;


    actualizarTextoResultados(
        equipos.length,
        inicioVisible,
        finVisible
    );


    /* =====================================
       RENDERIZAR FILAS
    ===================================== */

    equiposPagina.forEach(
        equipo => {

            const nombre =
                equipo.Nombre ||
                "Sin nombre";


            const estado =
                equipo.Estado ||
                "No verificable";


            const ip =
                equipo.IP ||
                "—";


            const sistemaOperativo =
                equipo.SistemaOperativo ||
                "—";


            const versionSO =
                equipo.VersionSO ||
                "";


            const area =
                equipo.Area ||
                "Sin identificar";


            const ultimoRegistro =
                formatearFechaEquipo(
                    equipo.UltimoRegistroAD
                );


            const tipoEstado =
                clasificarEstadoEquipo(
                    estado
                );


            const fila =
                document.createElement(
                    "tr"
                );


            fila.innerHTML = `

                <td>

                    <div class="equipo-name">

                        <div class="equipo-icon">
                            🖥️
                        </div>

                        <div>

                            <strong
                                title="${escaparHTML(nombre)}"
                            >
                                ${escaparHTML(nombre)}
                            </strong>

                            ${
                                equipo.DNSHostName
                                    ? `
                                        <span
                                            title="${escaparHTML(equipo.DNSHostName)}"
                                        >
                                            ${escaparHTML(equipo.DNSHostName)}
                                        </span>
                                    `
                                    : ""
                            }

                        </div>

                    </div>

                </td>


                <td>

                    <span
                        class="equipo-status ${tipoEstado}"
                    >

                        <span
                            class="equipo-status-dot"
                        ></span>

                        ${escaparHTML(estado)}

                    </span>

                </td>


                <td>

                    <span class="equipo-ip">
                        ${escaparHTML(ip)}
                    </span>

                </td>


                <td>

                    <span class="equipo-so">

                        ${escaparHTML(
                            sistemaOperativo
                        )}

                        ${
                            versionSO
                                ? `<br><small>${escaparHTML(versionSO)}</small>`
                                : ""
                        }

                    </span>

                </td>


                <td>

                    <span class="equipo-area">

                        ${escaparHTML(
                            area
                        )}

                    </span>

                </td>


                <td>

                    <span class="equipo-date">

                        ${escaparHTML(
                            ultimoRegistro
                        )}

                    </span>

                </td>

            `;


            tbody.appendChild(
                fila
            );

        }
    );


    /* =====================================
       PAGINACIÓN
    ===================================== */

    if (
        actualizarPaginacion
    ) {

        renderizarPaginacion(
            equipos.length
        );

    }

    else {

        renderizarPaginacion(
            equipos.length
        );

    }

}


/* =========================================
   ACTUALIZAR RESUMEN EQUIPOS
========================================= */

function actualizarResumenEquipos(
    equipos
) {

    const total =
        equipos.length;


    const conectados =
        equipos.filter(
            equipo =>
                String(
                    equipo.Estado || ""
                )
                .toLowerCase()
                .trim() ===
                "conectado"
        ).length;


    const desconectados =
        equipos.filter(
            equipo =>
                String(
                    equipo.Estado || ""
                )
                .toLowerCase()
                .trim() ===
                "desconectado"
        ).length;


    const totalElemento =
        document.getElementById(
            "totalEquipos"
        );


    const conectadosElemento =
        document.getElementById(
            "equiposConectados"
        );


    const desconectadosElemento =
        document.getElementById(
            "equiposDesconectados"
        );


    if (
        totalElemento
    ) {

        totalElemento.textContent =
            total;

    }


    if (
        conectadosElemento
    ) {

        conectadosElemento.textContent =
            conectados;

    }


    if (
        desconectadosElemento
    ) {

        desconectadosElemento.textContent =
            desconectados;

    }

}


/* =========================================
   CREAR FILTRO ÁREAS
========================================= */

function cargarFiltroAreas(
    equipos
) {

    const select =
        document.getElementById(
            "filtroAreaEquipo"
        );


    if (
        !select
    ) return;


    const areas =
        [
            ...new Set(
                equipos
                    .map(
                        equipo =>
                            equipo.Area ||
                            "Sin identificar"
                    )
                    .filter(
                        area =>
                            area
                    )
            )
        ]
        .sort(
            (a, b) =>
                String(a).localeCompare(
                    String(b),
                    "es"
                )
        );


    select.innerHTML = `

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


            select.appendChild(
                option
            );

        }
    );

}


/* =========================================
   APLICAR FILTROS
========================================= */

function aplicarFiltrosEquipos() {

    const buscar =
        document.getElementById(
            "buscarEquipo"
        );


    const filtroEstado =
        document.getElementById(
            "filtroEstadoEquipo"
        );


    const filtroArea =
        document.getElementById(
            "filtroAreaEquipo"
        );


    const texto =
        buscar
            ? buscar.value
                .trim()
                .toLowerCase()
            : "";


    const estadoSeleccionado =
        filtroEstado
            ? filtroEstado.value
            : "todos";


    const areaSeleccionada =
        filtroArea
            ? filtroArea.value
            : "todos";


    const filtrados =
        equiposData.filter(
            equipo => {

                const nombre =
                    String(
                        equipo.Nombre || ""
                    )
                    .toLowerCase();


                const ip =
                    String(
                        equipo.IP || ""
                    )
                    .toLowerCase();


                const dns =
                    String(
                        equipo.DNSHostName || ""
                    )
                    .toLowerCase();


                const area =
                    String(
                        equipo.Area ||
                        "Sin identificar"
                    );


                const estado =
                    String(
                        equipo.Estado || ""
                    );


                const coincideBusqueda =
                    !texto ||
                    nombre.includes(
                        texto
                    ) ||
                    ip.includes(
                        texto
                    ) ||
                    dns.includes(
                        texto
                    );


                const coincideEstado =
                    estadoSeleccionado ===
                        "todos" ||
                    estado ===
                        estadoSeleccionado;


                const coincideArea =
                    areaSeleccionada ===
                        "todos" ||
                    area ===
                        areaSeleccionada;


                return (
                    coincideBusqueda &&
                    coincideEstado &&
                    coincideArea
                );

            }
        );


    /*
     * Cada vez que cambia un filtro,
     * volvemos a la primera página.
     */

    paginaEquiposActual =
        1;


    renderizarEquipos(
        filtrados
    );

}


/* =========================================
   CONFIGURAR FILTROS
========================================= */

function configurarFiltrosEquipos() {

    const buscar =
        document.getElementById(
            "buscarEquipo"
        );


    const filtroEstado =
        document.getElementById(
            "filtroEstadoEquipo"
        );


    const filtroArea =
        document.getElementById(
            "filtroAreaEquipo"
        );


    if (
        buscar &&
        !buscar.dataset.configurado
    ) {

        buscar.addEventListener(
            "input",
            aplicarFiltrosEquipos
        );


        buscar.dataset.configurado =
            "true";

    }


    if (
        filtroEstado &&
        !filtroEstado.dataset.configurado
    ) {

        filtroEstado.addEventListener(
            "change",
            aplicarFiltrosEquipos
        );


        filtroEstado.dataset.configurado =
            "true";

    }


    if (
        filtroArea &&
        !filtroArea.dataset.configurado
    ) {

        filtroArea.addEventListener(
            "change",
            aplicarFiltrosEquipos
        );


        filtroArea.dataset.configurado =
            "true";

    }

}


/* =========================================
   CARGAR EQUIPOS
========================================= */

async function cargarEquipos() {

    try {

        console.log(
            "===================================="
        );

        console.log(
            "🖥️ INICIANDO INVENTARIO DE EQUIPOS"
        );

        console.log(
            "===================================="
        );


        const equipos =
            await obtenerEquipos();


        equiposData =
            equipos;


        actualizarResumenEquipos(
            equiposData
        );


        cargarFiltroAreas(
            equiposData
        );


        /*
         * Primera página.
         */

        paginaEquiposActual =
            1;


        renderizarEquipos(
            equiposData
        );


        configurarFiltrosEquipos();


        const estado =
            document.getElementById(
                "equiposEstado"
            );


        if (
            estado
        ) {

            estado.textContent =
                "● ACTUALIZADO";

            estado.className =
                "section-status online";

        }


        console.log(
            "✅ Inventario de equipos cargado correctamente."
        );


    }

    catch (
        error
    ) {

        console.error(
            "❌ No se pudo cargar el inventario:",
            error
        );


        const tbody =
            document.getElementById(
                "equiposTableBody"
            );


        if (
            tbody
        ) {

            tbody.innerHTML = `

                <tr>

                    <td
                        colspan="6"
                        class="equipos-loading"
                    >

                        ❌ No se pudo cargar
                        <strong>el inventario desde SharePoint</strong>.

                    </td>

                </tr>

            `;

        }


        const estado =
            document.getElementById(
                "equiposEstado"
            );


        if (
            estado
        ) {

            estado.textContent =
                "● ERROR";

            estado.className =
                "section-status offline";

        }

    }

}


/* =========================================
   CARGAR MONITOREO
========================================= */

async function cargarMonitoreo() {

    try {

        mostrarCarga(
            "Cargando monitoreo..."
        );


        /* =================================
           SHAREPOINT
        ================================= */

        const datos =
            await obtenerMonitoreoTI();


        renderizarServicios(
            datos
        );


        actualizarResumen(
            datos
        );


        actualizarEstadoGeneral(
            datos
        );


        /* =================================
           ÚLTIMA REVISIÓN
        ================================= */

        actualizarHora(
            datos
        );


        /* =================================
           EQUIPOS
        ================================= */

        cambiarMensaje(
            "Cargando inventario de equipos..."
        );


        await cargarEquipos();


        console.log(
            "✅ Centro de Monitoreo TI actualizado."
        );


    }

    catch (
        error
    ) {

        console.error(
            "❌ No se pudo cargar el monitoreo:",
            error
        );


        const elemento =
            document.getElementById(
                "estadoGeneral"
            );


        if (
            elemento
        ) {

            elemento.textContent =
                "● ERROR DE CONEXIÓN";

            elemento.className =
                "section-status offline";

        }


        const estadoMonitoreo =
            document.getElementById(
                "estadoMonitoreo"
            );


        if (
            estadoMonitoreo
        ) {

            estadoMonitoreo.textContent =
                "Error";

        }

    }

    finally {

        setTimeout(
            ocultarCarga,
            500
        );

    }

}


/* =========================================
   INICIO
========================================= */

document.addEventListener(
    "DOMContentLoaded",
    function () {

        console.log(
            "🚀 Iniciando Centro de Monitoreo TI..."
        );


        cargarMonitoreo();

    }
);


/* =========================================
   ACTUALIZACIÓN
   CADA 2 MINUTOS
========================================= */

setInterval(
    function () {

        console.log(
            "🔄 Actualización automática de monitoreo..."
        );


        cargarMonitoreo();

    },
    120000
);