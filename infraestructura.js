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

}

else {

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
   VARIABLES
========================================= */

let equiposData = [];

let paginaEquiposActual = 1;

let equiposFiltradosActuales = [];

const EQUIPOS_POR_PAGINA = 40;

let actualizacionEnCurso = false;


/* =========================================
   FETCH CON TIMEOUT
========================================= */

async function fetchConTimeout(
    url,
    opciones = {},
    tiempo = 20000
) {

    const controller =
        new AbortController();

    const timeout =
        setTimeout(
            () => controller.abort(),
            tiempo
        );

    try {

        return await fetch(
            url,
            {
                ...opciones,
                signal:
                    controller.signal
            }
        );

    }

    catch (error) {

        if (
            error.name === "AbortError"
        ) {

            throw new Error(
                "La solicitud tardó demasiado en responder."
            );

        }

        throw error;

    }

    finally {

        clearTimeout(
            timeout
        );

    }

}


/* =========================================
   MENSAJE DE CARGA
========================================= */

function cambiarMensaje(
    mensaje
) {

    const elemento =
        document.getElementById(
            "loadingMessage"
        );

    if (
        elemento
    ) {

        elemento.textContent =
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

    const overlay =
        document.getElementById(
            "loadingOverlay"
        );

    if (
        overlay
    ) {

        overlay.classList.remove(
            "hidden"
        );

        overlay.classList.remove(
            "oculto"
        );

    }

}


/* =========================================
   OCULTAR CARGA
========================================= */

function ocultarCarga() {

    const overlay =
        document.getElementById(
            "loadingOverlay"
        );

    if (
        overlay
    ) {

        overlay.classList.add(
            "hidden"
        );

        overlay.classList.add(
            "oculto"
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
            "❌ Error obteniendo token:",
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

    const url =
        `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOSTNAME}:${SHAREPOINT_SITE}`;

    const respuesta =
        await fetchConTimeout(
            url,
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
        !respuesta.ok
    ) {

        const error =
            await respuesta.text();

        console.error(
            "❌ Error SharePoint:",
            respuesta.status,
            error
        );

        throw new Error(
            `SharePoint HTTP ${respuesta.status}`
        );

    }

    const sitio =
        await respuesta.json();

    console.log(
        "✅ SharePoint encontrado:",
        sitio.id
    );

    return sitio;

}


/* =========================================
   OBTENER MONITOREO TI
========================================= */

async function obtenerMonitoreoTI() {

    console.log(
        "☁️ Cargando MonitoreoTI..."
    );

    const TOKEN =
        await obtenerTokenInfraestructura();

    const sitio =
        await obtenerSitioSharePoint(
            TOKEN
        );

    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${SHAREPOINT_LIST}/items?expand=fields&$top=999`;

    const respuesta =
        await fetchConTimeout(
            url,
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
        !respuesta.ok
    ) {

        const error =
            await respuesta.text();

        console.error(
            "❌ Error MonitoreoTI:",
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

async function obtenerEquipos() {

    console.log(
        "🖥️ Cargando equipos.json..."
    );

    const TOKEN =
        await obtenerTokenInfraestructura();

    const sitio =
        await obtenerSitioSharePoint(
            TOKEN
        );

    const respuestaDrive =
        await fetchConTimeout(

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

        const error =
            await respuestaDrive.text();

        console.error(
            "❌ Error Drive:",
            respuestaDrive.status,
            error
        );

        throw new Error(
            `Drive HTTP ${respuestaDrive.status}`
        );

    }

    const drive =
        await respuestaDrive.json();

    if (
        !drive.id
    ) {

        throw new Error(
            "El Drive no tiene ID."
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
        await fetchConTimeout(

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

        const error =
            await respuestaArchivo.text();

        console.error(
            "❌ Error equipos.json:",
            respuestaArchivo.status,
            error
        );

        throw new Error(
            `equipos.json HTTP ${respuestaArchivo.status}`
        );

    }

    const datos =
        await respuestaArchivo.json();

    if (
        !Array.isArray(datos)
    ) {

        throw new Error(
            "equipos.json no contiene un array."
        );

    }

    console.log(
        "✅ Equipos encontrados:",
        datos.length
    );

    return datos;

}


/* =========================================
   OBTENER CAMPOS
========================================= */

function obtenerCampos(
    item
) {

    return item?.fields || {};

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
   ICONO SERVICIO
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
   FORMATEAR FECHA
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
   ACTUALIZAR HORA
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

    const fechas =
        datos
            .map(
                item =>
                    obtenerCampos(
                        item
                    ).Ultimarevision
            )
            .filter(
                Boolean
            );

    if (
        !fechas.length
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
        !fechasValidas.length
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

            estados.push(
                normalizarEstado(
                    estado
                )
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

}


/* =========================================
   ACTUALIZAR RESUMEN
========================================= */

function actualizarResumen(
    datos
) {

    let activos = 0;
    let advertencias = 0;
    let incidencias = 0;

    datos.forEach(
        item => {

            const estado =
                clasificarEstado(
                    obtenerCampos(
                        item
                    ).Estado
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
   ESTADO GENERAL
========================================= */

function actualizarEstadoGeneral(
    datos
) {

    const elemento =
        document.getElementById(
            "estadoMonitoreo"
        );

    if (
        !elemento
    ) return;

    if (
        !datos ||
        datos.length === 0
    ) {

        elemento.textContent =
            "No configurado";

        elemento.className =
            "status-badge neutral";

        return;

    }

    const estados =
        datos.map(
            item =>
                normalizarEstado(
                    obtenerCampos(
                        item
                    ).Estado
                )
                .toLowerCase()
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

        elemento.textContent =
            "Incidencia";

        elemento.className =
            "status-badge offline";

    }

    else if (
        tieneAdvertencia
    ) {

        elemento.textContent =
            "Advertencia";

        elemento.className =
            "status-badge warning";

    }

    else if (
        todosOperativos
    ) {

        elemento.textContent =
            "Operativo";

        elemento.className =
            "status-badge online";

    }

    else {

        elemento.textContent =
            "No configurado";

        elemento.className =
            "status-badge neutral";

    }

}


/* =========================================
   ESTADO EQUIPO
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
   FECHA EQUIPO
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
   RENDERIZAR EQUIPOS
========================================= */

function renderizarEquipos(
    equipos
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

    equiposFiltradosActuales =
        equipos || [];

    tbody.innerHTML =
        "";

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

        const resultados =
            document.getElementById(
                "equiposResultados"
            );

        if (
            resultados
        ) {

            resultados.textContent =
                "0 equipos";

        }

        return;

    }

    if (
        empty
    ) {

        empty.style.display =
            "none";

    }

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

    const resultados =
        document.getElementById(
            "equiposResultados"
        );

    if (
        resultados
    ) {

        resultados.textContent =
            `Mostrando ${inicioIndice + 1}–${finIndice} de ${equipos.length} equipos`;

    }

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

                    <span class="trabajador-sin-asignar">
                        —
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

}


/* =========================================
   RESUMEN EQUIPOS
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
                .trim()
                .toLowerCase() ===
                "conectado"
        ).length;

    const desconectados =
        equipos.filter(
            equipo =>
                String(
                    equipo.Estado || ""
                )
                .trim()
                .toLowerCase() ===
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
   FILTRO ÁREAS
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
   FILTROS
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
   PAGINACIÓN
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

    const anterior =
        document.createElement(
            "button"
        );

    anterior.type =
        "button";

    anterior.className =
        "pagination-btn pagination-prev";

    anterior.textContent =
        "‹";

    anterior.disabled =
        paginaEquiposActual === 1;

    anterior.addEventListener(
        "click",
        () => {

            if (
                paginaEquiposActual > 1
            ) {

                paginaEquiposActual--;

                renderizarEquipos(
                    equiposFiltradosActuales
                );

            }

        }
    );

    contenedor.appendChild(
        anterior
    );

    for (
        let pagina = 1;
        pagina <= totalPaginas;
        pagina++
    ) {

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
            () => {

                paginaEquiposActual =
                    pagina;

                renderizarEquipos(
                    equiposFiltradosActuales
                );

            }
        );

        contenedor.appendChild(
            boton
        );

    }

    const siguiente =
        document.createElement(
            "button"
        );

    siguiente.type =
        "button";

    siguiente.className =
        "pagination-btn pagination-next";

    siguiente.textContent =
        "›";

    siguiente.disabled =
        paginaEquiposActual ===
        totalPaginas;

    siguiente.addEventListener(
        "click",
        () => {

            if (
                paginaEquiposActual <
                totalPaginas
            ) {

                paginaEquiposActual++;

                renderizarEquipos(
                    equiposFiltradosActuales
                );

            }

        }
    );

    contenedor.appendChild(
        siguiente
    );

}


/* =========================================
   CARGAR EQUIPOS
========================================= */

async function cargarEquipos() {

    try {

        console.log(
            "🖥️ INICIANDO INVENTARIO..."
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
            "✅ INVENTARIO CARGADO:",
            equiposData.length
        );

        return true;

    }

    catch (
        error
    ) {

        console.error(
            "❌ ERROR INVENTARIO:",
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
                        colspan="7"
                        class="equipos-loading"
                    >

                        ❌ No se pudo cargar
                        <strong>
                            el inventario desde SharePoint
                        </strong>.

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

        return false;

    }

}


/* =========================================
   CARGAR SERVICIOS
========================================= */

async function cargarServicios() {

    try {

        console.log(
            "☁️ INICIANDO SERVICIOS..."
        );

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

        actualizarHora(
            datos
        );

        console.log(
            "✅ SERVICIOS CARGADOS:",
            datos.length
        );

        return true;

    }

    catch (
        error
    ) {

        console.error(
            "❌ ERROR SERVICIOS:",
            error
        );

        const grid =
            document.getElementById(
                "serviciosGrid"
            );

        if (
            grid
        ) {

            grid.innerHTML = `

                <div class="empty-state">

                    ❌ No se pudieron cargar
                    los servicios.

                </div>

            `;

        }

        return false;

    }

}


/* =========================================
   CARGAR MONITOREO
========================================= */

async function cargarMonitoreo() {

    if (
        actualizacionEnCurso
    ) {

        console.log(
            "⏳ Ya hay una actualización ejecutándose."
        );

        return;

    }

    actualizacionEnCurso =
        true;

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
        "Cargando monitoreo..."
    );

    try {

        /*
         * IMPORTANTE:
         *
         * No hacemos que equipos dependa
         * de MonitoreoTI.
         *
         * Cada sección intenta cargar
         * independientemente.
         */

        cambiarMensaje(
            "Cargando servicios..."
        );

        await cargarServicios();


        cambiarMensaje(
            "Cargando inventario de equipos..."
        );

        await cargarEquipos();


        console.log(
            "===================================="
        );

        console.log(
            "✅ CENTRO DE MONITOREO CARGADO"
        );

        console.log(
            "===================================="
        );

    }

    catch (
        error
    ) {

        console.error(
            "❌ ERROR GENERAL:",
            error
        );

    }

    finally {

        /*
         * Nunca dejar el overlay bloqueando
         * la página.
         */

        ocultarCarga();

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
            "🚀 DOM LISTO"
        );

        cargarMonitoreo();

    }
);


/* =========================================
   ACTUALIZACIÓN CADA 2 MINUTOS
========================================= */

setInterval(
    function () {

        console.log(
            "🔄 ACTUALIZACIÓN AUTOMÁTICA"
        );

        cargarMonitoreo();

    },
    120000
);