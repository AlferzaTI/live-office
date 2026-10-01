/* ============================================================
   ALFERZA LIVE OFFICE
   INFRAESTRUCTURA TI
   ============================================================ */


/* ============================================================
   CONFIGURACIÓN MICROSOFT / MSAL
   ============================================================ */

const msalConfigInfraestructura = {
    auth: {
        clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",

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


/* ============================================================
   SHAREPOINT
   ============================================================ */

const SHAREPOINT_HOSTNAME =
    "alferzaholding-my.sharepoint.com";

const SHAREPOINT_SITE =
    "/personal/soporte1_alferza_pe";

const SHAREPOINT_LIST =
    "MonitoreoTI";

const EQUIPOS_PATH =
    "Procedimientos T.I/equipos.json";


/* ============================================================
   VARIABLES
   ============================================================ */

let msalInstanceInfraestructura = null;

let equiposData = [];

let equiposFiltrados = [];

let paginaActual = 1;

const equiposPorPagina = 20;

let actualizacionEnCurso = false;


/* ============================================================
   ELEMENTOS HTML
   ============================================================ */

const loadingOverlay =
    document.getElementById("loadingOverlay");

const loadingMessage =
    document.getElementById("loadingMessage");

const serviciosContainer =
    document.getElementById("serviciosContainer");

const equiposTableBody =
    document.getElementById("equiposTableBody");

const equiposResultados =
    document.getElementById("equiposResultados");

const totalEquipos =
    document.getElementById("totalEquipos");

const equiposConectados =
    document.getElementById("equiposConectados");

const equiposDesconectados =
    document.getElementById("equiposDesconectados");

const buscarEquipo =
    document.getElementById("buscarEquipo");

const filtroEstadoEquipo =
    document.getElementById("filtroEstadoEquipo");

const filtroAreaEquipo =
    document.getElementById("filtroAreaEquipo");

const equiposPagination =
    document.getElementById("equiposPagination");

const estadoGeneral =
    document.getElementById("estadoGeneral");

const ultimaActualizacion =
    document.getElementById("ultimaActualizacion");


/* ============================================================
   LOG
   ============================================================ */

function log(...mensajes) {
    console.log(
        "[ALFERZA INFRA]",
        ...mensajes
    );
}


/* ============================================================
   LOADING
   ============================================================ */

function cambiarMensaje(mensaje) {

    if (loadingMessage) {
        loadingMessage.textContent = mensaje;
    }
}


function mostrarCarga(mensaje = "Cargando...") {

    cambiarMensaje(mensaje);

    if (!loadingOverlay) {
        return;
    }

    loadingOverlay.classList.remove("hidden");
    loadingOverlay.classList.remove("oculto");

    loadingOverlay.style.display = "flex";
}


function ocultarCarga() {

    if (!loadingOverlay) {
        return;
    }

    loadingOverlay.classList.add("hidden");
    loadingOverlay.classList.add("oculto");

    loadingOverlay.style.display = "none";
}


/* ============================================================
   TIMEOUT PARA FETCH
   ============================================================ */

async function fetchConTimeout(
    url,
    opciones = {},
    timeout = 20000
) {

    const controller =
        new AbortController();

    const temporizador =
        setTimeout(
            () => controller.abort(),
            timeout
        );

    try {

        const respuesta =
            await fetch(
                url,
                {
                    ...opciones,
                    signal: controller.signal
                }
            );

        return respuesta;

    } catch (error) {

        if (error.name === "AbortError") {

            throw new Error(
                "La solicitud tardó demasiado en responder."
            );
        }

        throw error;

    } finally {

        clearTimeout(temporizador);
    }
}


/* ============================================================
   INICIALIZAR MSAL
   ============================================================ */

function inicializarMSAL() {

    if (typeof msal === "undefined") {

        console.error(
            "❌ MSAL no está disponible."
        );

        return false;
    }

    try {

        msalInstanceInfraestructura =
            new msal.PublicClientApplication(
                msalConfigInfraestructura
            );

        log(
            "✅ MSAL inicializado."
        );

        return true;

    } catch (error) {

        console.error(
            "❌ Error inicializando MSAL:",
            error
        );

        return false;
    }
}


/* ============================================================
   OBTENER CUENTA MICROSOFT
   ============================================================ */

async function obtenerCuentaMicrosoft() {

    if (!msalInstanceInfraestructura) {

        throw new Error(
            "MSAL no está inicializado."
        );
    }


    /*
     * Procesar respuesta de login si existe.
     */
    try {

        const respuestaRedirect =
            await msalInstanceInfraestructura
                .handleRedirectPromise();

        if (respuestaRedirect?.account) {

            msalInstanceInfraestructura
                .setActiveAccount(
                    respuestaRedirect.account
                );

            log(
                "✅ Login recibido:",
                respuestaRedirect.account.username
            );
        }

    } catch (error) {

        console.error(
            "❌ Error procesando login:",
            error
        );
    }


    let cuenta =
        msalInstanceInfraestructura
            .getActiveAccount();


    if (!cuenta) {

        const cuentas =
            msalInstanceInfraestructura
                .getAllAccounts();


        if (cuentas.length > 0) {

            cuenta =
                cuentas[0];

            msalInstanceInfraestructura
                .setActiveAccount(
                    cuenta
                );
        }
    }


    return cuenta;
}


/* ============================================================
   OBTENER TOKEN
   ============================================================ */

async function obtenerTokenInfraestructura() {

    log(
        "🔐 Obteniendo token Microsoft..."
    );


    const cuenta =
        await obtenerCuentaMicrosoft();


    /*
     * Si no hay sesión, iniciar login.
     */
    if (!cuenta) {

        log(
            "⚠️ No existe sesión. Iniciando login..."
        );

        cambiarMensaje(
            "Iniciando sesión Microsoft..."
        );


        await msalInstanceInfraestructura
            .loginRedirect({

                scopes: [
                    "User.Read",
                    "Sites.Read.All"
                ]
            });


        /*
         * loginRedirect redirige la página.
         */
        return null;
    }


    log(
        "👤 Cuenta:",
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

                    account: cuenta
                });


        log(
            "✅ Token obtenido."
        );


        return respuesta.accessToken;

    } catch (error) {

        console.error(
            "❌ No se pudo obtener token silencioso:",
            error
        );


        /*
         * Si requiere interacción,
         * volver a iniciar sesión.
         */
        if (
            error instanceof
            msal.InteractionRequiredAuthError
        ) {

            log(
                "⚠️ Se requiere autenticación."
            );


            await msalInstanceInfraestructura
                .loginRedirect({

                    scopes: [
                        "User.Read",
                        "Sites.Read.All"
                    ]
                });


            return null;
        }


        throw error;
    }
}


/* ============================================================
   OBTENER SITIO SHAREPOINT
   ============================================================ */

async function obtenerSitioSharePoint(TOKEN) {

    log(
        "🔎 Buscando sitio SharePoint..."
    );


    const url =
        `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOSTNAME}:${SHAREPOINT_SITE}`;


    const respuesta =
        await fetchConTimeout(

            url,

            {
                method: "GET",

                headers: {
                    Authorization:
                        `Bearer ${TOKEN}`
                }
            },

            20000
        );


    log(
        "📡 SharePoint:",
        respuesta.status
    );


    if (!respuesta.ok) {

        const texto =
            await respuesta.text();


        console.error(
            "❌ Error obteniendo sitio:",
            texto
        );


        throw new Error(
            `No se pudo acceder al sitio SharePoint (${respuesta.status}).`
        );
    }


    const sitio =
        await respuesta.json();


    log(
        "✅ Sitio encontrado:",
        sitio.id
    );


    return sitio;
}


/* ============================================================
   OBTENER MONITOREO TI
   ============================================================ */

async function obtenerMonitoreoTI(
    TOKEN,
    sitio
) {

    log(
        "📊 Cargando MonitoreoTI..."
    );


    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${SHAREPOINT_LIST}/items?expand=fields`;


    const respuesta =
        await fetchConTimeout(

            url,

            {
                method: "GET",

                headers: {
                    Authorization:
                        `Bearer ${TOKEN}`
                }
            },

            20000
        );


    log(
        "📡 MonitoreoTI:",
        respuesta.status
    );


    if (!respuesta.ok) {

        const texto =
            await respuesta.text();


        console.error(
            "❌ Error MonitoreoTI:",
            texto
        );


        throw new Error(
            `No se pudo cargar MonitoreoTI (${respuesta.status}).`
        );
    }


    const datos =
        await respuesta.json();


    log(
        "✅ Registros MonitoreoTI:",
        datos.value?.length || 0
    );


    return datos.value || [];
}


/* ============================================================
   OBTENER EQUIPOS.JSON
   ============================================================ */

async function obtenerEquipos(
    TOKEN,
    sitio
) {

    log(
        "🖥️ Cargando equipos.json..."
    );


    /*
     * Primero obtenemos el Drive del sitio.
     */
    const driveUrl =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/drive`;


    const driveResponse =
        await fetchConTimeout(

            driveUrl,

            {
                method: "GET",

                headers: {
                    Authorization:
                        `Bearer ${TOKEN}`
                }
            },

            20000
        );


    log(
        "📁 Drive:",
        driveResponse.status
    );


    if (!driveResponse.ok) {

        const texto =
            await driveResponse.text();


        console.error(
            "❌ Error obteniendo Drive:",
            texto
        );


        throw new Error(
            `No se pudo acceder al almacenamiento de SharePoint (${driveResponse.status}).`
        );
    }


    const drive =
        await driveResponse.json();


    log(
        "✅ Drive:",
        drive.id
    );


    /*
     * IMPORTANTE:
     *
     * No usamos encodeURIComponent() sobre
     * toda la ruta porque convertiría "/" en "%2F".
     */
    const rutaArchivo =
        EQUIPOS_PATH
            .split("/")
            .map(
                parte =>
                    encodeURIComponent(parte)
            )
            .join("/");


    const archivoUrl =
        `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${rutaArchivo}:/content`;


    log(
        "📄 Archivo:",
        EQUIPOS_PATH
    );


    const archivoResponse =
        await fetchConTimeout(

            archivoUrl,

            {
                method: "GET",

                headers: {
                    Authorization:
                        `Bearer ${TOKEN}`
                }
            },

            30000
        );


    log(
        "📡 equipos.json:",
        archivoResponse.status
    );


    if (!archivoResponse.ok) {

        const texto =
            await archivoResponse.text();


        console.error(
            "❌ Error descargando equipos.json:",
            texto
        );


        throw new Error(
            `No se pudo descargar equipos.json (${archivoResponse.status}).`
        );
    }


    const texto =
        await archivoResponse.text();


    let datos;


    try {

        datos =
            JSON.parse(texto);

    } catch (error) {

        console.error(
            "❌ Error leyendo JSON:",
            error
        );


        throw new Error(
            "El archivo equipos.json no contiene un JSON válido."
        );
    }


    if (!Array.isArray(datos)) {

        throw new Error(
            "El contenido de equipos.json no es una lista."
        );
    }


    log(
        "✅ Equipos encontrados:",
        datos.length
    );


    return datos;
}


/* ============================================================
   RENDERIZAR SERVICIOS
   ============================================================ */

function renderizarServicios(datos) {

    if (!serviciosContainer) {
        return;
    }


    if (!datos || datos.length === 0) {

        serviciosContainer.innerHTML = `
            <div class="service-loading">
                No hay servicios registrados.
            </div>
        `;

        return;
    }


    serviciosContainer.innerHTML =
        datos
            .map(item => {

                const fields =
                    item.fields || item;


                const nombre =
                    fields.Title ||
                    fields.Nombre ||
                    fields.Servicio ||
                    fields.Name ||
                    "Servicio";


                const estado =
                    fields.Estado ||
                    fields.Status ||
                    fields.EstadoServicio ||
                    "Desconocido";


                const estadoTexto =
                    String(estado);


                const estadoNormalizado =
                    estadoTexto
                        .toLowerCase()
                        .normalize("NFD")
                        .replace(
                            /[\u0300-\u036f]/g,
                            ""
                        );


                let clase =
                    "unknown";


                if (

                    estadoNormalizado.includes(
                        "activo"
                    ) ||

                    estadoNormalizado.includes(
                        "operativo"
                    ) ||

                    estadoNormalizado.includes(
                        "conectado"
                    ) ||

                    estadoNormalizado.includes(
                        "online"
                    ) ||

                    estadoNormalizado.includes(
                        "ok"
                    )

                ) {

                    clase =
                        "online";

                } else if (

                    estadoNormalizado.includes(
                        "inactivo"
                    ) ||

                    estadoNormalizado.includes(
                        "desconectado"
                    ) ||

                    estadoNormalizado.includes(
                        "offline"
                    ) ||

                    estadoNormalizado.includes(
                        "error"
                    ) ||

                    estadoNormalizado.includes(
                        "caido"
                    )

                ) {

                    clase =
                        "offline";
                }


                return `

                    <div class="service-card">

                        <div class="service-card-header">

                            <div class="service-name">
                                ${escaparHTML(nombre)}
                            </div>

                            <span
                                class="service-status ${clase}"
                            >
                                ${escaparHTML(estadoTexto)}
                            </span>

                        </div>

                    </div>

                `;

            })
            .join("");
}


/* ============================================================
   RESUMEN DE EQUIPOS
   ============================================================ */

function actualizarResumenEquipos(datos) {

    const total =
        datos.length;


    const conectados =
        datos.filter(
            equipo =>
                normalizarEstado(
                    equipo.Estado
                ) === "conectado"
        ).length;


    const desconectados =
        datos.filter(
            equipo =>
                normalizarEstado(
                    equipo.Estado
                ) === "desconectado"
        ).length;


    if (totalEquipos) {

        totalEquipos.textContent =
            total;
    }


    if (equiposConectados) {

        equiposConectados.textContent =
            conectados;
    }


    if (equiposDesconectados) {

        equiposDesconectados.textContent =
            desconectados;
    }
}


/* ============================================================
   NORMALIZAR ESTADO
   ============================================================ */

function normalizarEstado(valor) {

    return String(
        valor || ""
    )
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        );
}


/* ============================================================
   CARGAR ÁREAS
   ============================================================ */

function cargarAreas() {

    if (!filtroAreaEquipo) {
        return;
    }


    const areas =
        [
            ...new Set(

                equiposData
                    .map(
                        equipo =>
                            equipo.Area
                    )
                    .filter(
                        area =>
                            area &&
                            String(area).trim()
                    )
            )
        ]
        .sort(
            (a, b) =>
                String(a)
                    .localeCompare(
                        String(b)
                    )
        );


    filtroAreaEquipo.innerHTML = `

        <option value="Todos">
            Todas
        </option>

    `;


    areas.forEach(area => {

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
    });
}


/* ============================================================
   APLICAR FILTROS
   ============================================================ */

function aplicarFiltrosEquipos() {

    const texto =
        String(
            buscarEquipo?.value || ""
        )
            .trim()
            .toLowerCase();


    const estado =
        filtroEstadoEquipo?.value ||
        "Todos";


    const area =
        filtroAreaEquipo?.value ||
        "Todos";


    equiposFiltrados =
        equiposData.filter(
            equipo => {

                const textoEquipo = [

                    equipo.Nombre,

                    equipo.IP,

                    equipo.SistemaOperativo,

                    equipo.VersionSO,

                    equipo.Area,

                    equipo.DNSHostName,

                    equipo.UltimoRegistroAD

                ]
                    .map(
                        valor =>
                            String(
                                valor || ""
                            )
                                .toLowerCase()
                    )
                    .join(" ");


                const coincideTexto =
                    !texto ||
                    textoEquipo.includes(
                        texto
                    );


                const coincideEstado =
                    estado === "Todos" ||
                    String(
                        equipo.Estado ||
                        ""
                    ) === estado;


                const coincideArea =
                    area === "Todos" ||
                    String(
                        equipo.Area ||
                        ""
                    ) === area;


                return (
                    coincideTexto &&
                    coincideEstado &&
                    coincideArea
                );
            }
        );


    paginaActual = 1;


    renderizarEquipos();
}


/* ============================================================
   RENDERIZAR EQUIPOS
   ============================================================ */

function renderizarEquipos() {

    if (!equiposTableBody) {
        return;
    }


    if (!equiposFiltrados.length) {

        equiposTableBody.innerHTML = `

            <tr>

                <td
                    colspan="6"
                    class="table-empty"
                >
                    No se encontraron equipos.
                </td>

            </tr>

        `;


        if (equiposResultados) {

            equiposResultados.textContent =
                "0";
        }


        renderizarPaginacion();

        return;
    }


    const total =
        equiposFiltrados.length;


    const inicio =
        (paginaActual - 1) *
        equiposPorPagina;


    const fin =
        inicio +
        equiposPorPagina;


    const equiposPagina =
        equiposFiltrados.slice(
            inicio,
            fin
        );


    equiposTableBody.innerHTML =
        equiposPagina
            .map(
                equipo => {

                    const estado =
                        equipo.Estado ||
                        "No verificable";


                    const estadoNormalizado =
                        normalizarEstado(
                            estado
                        );


                    let claseEstado =
                        "unknown";


                    if (
                        estadoNormalizado ===
                        "conectado"
                    ) {

                        claseEstado =
                            "connected";

                    } else if (
                        estadoNormalizado ===
                        "desconectado"
                    ) {

                        claseEstado =
                            "disconnected";

                    } else if (
                        estadoNormalizado ===
                        "no verificable"
                    ) {

                        claseEstado =
                            "unknown";
                    }


                    return `

                        <tr>

                            <td>
                                <strong>
                                    ${escaparHTML(
                                        equipo.Nombre ||
                                        "—"
                                    )}
                                </strong>
                            </td>


                            <td>

                                <span
                                    class="equipment-status ${claseEstado}"
                                >
                                    ${escaparHTML(
                                        estado
                                    )}
                                </span>

                            </td>


                            <td>
                                ${escaparHTML(
                                    equipo.IP ||
                                    "—"
                                )}
                            </td>


                            <td>
                                ${escaparHTML(
                                    equipo.SistemaOperativo ||
                                    "—"
                                )}
                            </td>


                            <td>
                                ${escaparHTML(
                                    equipo.Area ||
                                    "—"
                                )}
                            </td>


                            <td>
                                ${escaparHTML(
                                    equipo.UltimoRegistroAD ||
                                    "—"
                                )}
                            </td>

                        </tr>

                    `;
                }
            )
            .join("");


    if (equiposResultados) {

        equiposResultados.textContent =
            total;
    }


    renderizarPaginacion();
}


/* ============================================================
   PAGINACIÓN
   ============================================================ */

function renderizarPaginacion() {

    if (!equiposPagination) {
        return;
    }


    const totalPaginas =
        Math.ceil(
            equiposFiltrados.length /
            equiposPorPagina
        );


    if (totalPaginas <= 1) {

        equiposPagination.innerHTML =
            "";

        return;
    }


    let html = "";


    /*
     * ANTERIOR
     */

    html += `

        <button
            class="pagination-button"
            ${paginaActual <= 1 ? "disabled" : ""}
            onclick="cambiarPagina(${paginaActual - 1})"
        >
            ←
        </button>

    `;


    /*
     * NÚMEROS
     */

    for (
        let i = 1;
        i <= totalPaginas;
        i++
    ) {

        html += `

            <button
                class="pagination-button ${
                    i === paginaActual
                        ? "active"
                        : ""
                }"
                onclick="cambiarPagina(${i})"
            >
                ${i}
            </button>

        `;
    }


    /*
     * SIGUIENTE
     */

    html += `

        <button
            class="pagination-button"
            ${
                paginaActual >= totalPaginas
                    ? "disabled"
                    : ""
            }
            onclick="cambiarPagina(${paginaActual + 1})"
        >
            →
        </button>

    `;


    equiposPagination.innerHTML =
        html;
}


function cambiarPagina(pagina) {

    const totalPaginas =
        Math.ceil(
            equiposFiltrados.length /
            equiposPorPagina
        );


    if (
        pagina < 1 ||
        pagina > totalPaginas
    ) {

        return;
    }


    paginaActual =
        pagina;


    renderizarEquipos();
}


/* ============================================================
   ACTUALIZAR ESTADO GENERAL
   ============================================================ */

function actualizarEstadoGeneral(datos) {

    if (!estadoGeneral) {
        return;
    }


    if (!datos || !datos.length) {

        estadoGeneral.textContent =
            "Sin información";

        return;
    }


    let correctos = 0;


    datos.forEach(item => {

        const fields =
            item.fields || item;


        const estado =
            normalizarEstado(

                fields.Estado ||
                fields.Status ||
                fields.EstadoServicio
            );


        if (

            estado.includes(
                "activo"
            ) ||

            estado.includes(
                "operativo"
            ) ||

            estado.includes(
                "conectado"
            ) ||

            estado.includes(
                "online"
            ) ||

            estado === "ok"

        ) {

            correctos++;
        }
    });


    if (
        correctos === datos.length
    ) {

        estadoGeneral.textContent =
            "Todos los servicios operativos";

    } else if (
        correctos > 0
    ) {

        estadoGeneral.textContent =
            "Algunos servicios requieren atención";

    } else {

        estadoGeneral.textContent =
            "Se detectaron servicios con problemas";
    }
}


/* ============================================================
   ACTUALIZAR HORA
   ============================================================ */

function actualizarHora(datos) {

    if (!ultimaActualizacion) {
        return;
    }


    let fecha = null;


    /*
     * Buscar fecha dentro de MonitoreoTI.
     */

    if (datos && datos.length) {

        for (
            const item of datos
        ) {

            const fields =
                item.fields || item;


            const valorFecha =

                fields.Fecha ||
                fields.FechaActualizacion ||
                fields.UltimaActualizacion ||
                fields.Modified;


            if (!valorFecha) {
                continue;
            }


            const posibleFecha =
                new Date(
                    valorFecha
                );


            if (
                !isNaN(
                    posibleFecha.getTime()
                )
            ) {

                fecha =
                    posibleFecha;

                break;
            }
        }
    }


    /*
     * Si no existe fecha,
     * usar la hora actual.
     */

    if (!fecha) {

        fecha =
            new Date();
    }


    ultimaActualizacion.textContent =
        fecha.toLocaleString(
            "es-PE",
            {
                dateStyle: "short",
                timeStyle: "medium"
            }
        );
}


/* ============================================================
   CARGAR EQUIPOS
   ============================================================ */

async function cargarEquipos(
    TOKEN,
    sitio
) {

    try {

        log(
            "🖥️ Iniciando carga de equipos..."
        );


        const datos =
            await obtenerEquipos(
                TOKEN,
                sitio
            );


        equiposData =
            datos;


        equiposFiltrados =
            [...datos];


        actualizarResumenEquipos(
            equiposData
        );


        cargarAreas();


        renderizarEquipos();


        log(
            "✅ Inventario cargado:",
            datos.length
        );


        return true;

    } catch (error) {

        console.error(
            "❌ Error cargando equipos:",
            error
        );


        equiposData = [];

        equiposFiltrados = [];


        actualizarResumenEquipos([]);


        if (equiposTableBody) {

            equiposTableBody.innerHTML = `

                <tr>

                    <td
                        colspan="6"
                        class="table-error"
                    >

                        Error cargando inventario:
                        ${escaparHTML(
                            error.message
                        )}

                    </td>

                </tr>

            `;
        }


        if (equiposResultados) {

            equiposResultados.textContent =
                "0";
        }


        return false;
    }
}


/* ============================================================
   CARGAR SERVICIOS
   ============================================================ */

async function cargarServicios(
    TOKEN,
    sitio
) {

    try {

        log(
            "📊 Iniciando carga de servicios..."
        );


        const datos =
            await obtenerMonitoreoTI(
                TOKEN,
                sitio
            );


        renderizarServicios(
            datos
        );


        actualizarEstadoGeneral(
            datos
        );


        actualizarHora(
            datos
        );


        log(
            "✅ Servicios cargados."
        );


        return true;

    } catch (error) {

        console.error(
            "❌ Error cargando servicios:",
            error
        );


        if (serviciosContainer) {

            serviciosContainer.innerHTML = `

                <div class="service-loading">

                    No se pudieron cargar los servicios.

                    <br><br>

                    <small>
                        ${escaparHTML(
                            error.message
                        )}
                    </small>

                </div>

            `;
        }


        if (estadoGeneral) {

            estadoGeneral.textContent =
                "No se pudo verificar el estado";
        }


        return false;
    }
}


/* ============================================================
   DASHBOARD PRINCIPAL
   ============================================================ */

async function cargarDashboard() {

    if (actualizacionEnCurso) {

        log(
            "⏳ Ya hay una actualización en curso."
        );

        return;
    }


    actualizacionEnCurso =
        true;


    try {

        mostrarCarga(
            "Iniciando..."
        );


        log(
            "======================================"
        );

        log(
            "🚀 INICIANDO INFRAESTRUCTURA TI"
        );

        log(
            "======================================"
        );


        /*
         * 1. MSAL
         */

        cambiarMensaje(
            "Inicializando Microsoft..."
        );


        if (!msalInstanceInfraestructura) {

            const correcto =
                inicializarMSAL();


            if (!correcto) {

                throw new Error(
                    "No se pudo inicializar Microsoft Authentication."
                );
            }
        }


        /*
         * 2. TOKEN
         */

        cambiarMensaje(
            "Verificando sesión..."
        );


        const TOKEN =
            await obtenerTokenInfraestructura();


        /*
         * Si loginRedirect fue ejecutado,
         * la página se va a redirigir.
         */

        if (!TOKEN) {

            return;
        }


        /*
         * 3. SHAREPOINT
         */

        cambiarMensaje(
            "Conectando con SharePoint..."
        );


        const sitio =
            await obtenerSitioSharePoint(
                TOKEN
            );


        /*
         * 4. CARGAR SERVICIOS Y EQUIPOS
         *
         * Se ejecutan independientemente.
         *
         * Si MonitoreoTI falla,
         * equipos.json todavía puede cargar.
         */

        cambiarMensaje(
            "Cargando información..."
        );


        const resultados =
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


        log(
            "📊 Resultado servicios:",
            resultados[0].status
        );


        log(
            "🖥️ Resultado equipos:",
            resultados[1].status
        );


        log(
            "======================================"
        );

        log(
            "✅ CARGA FINALIZADA"
        );

        log(
            "======================================"
        );


    } catch (error) {

        console.error(
            "❌ ERROR GENERAL:",
            error
        );


        if (estadoGeneral) {

            estadoGeneral.textContent =
                "Error de conexión";
        }


        if (equiposTableBody) {

            equiposTableBody.innerHTML = `

                <tr>

                    <td
                        colspan="6"
                        class="table-error"
                    >

                        No se pudo cargar el inventario.

                        <br>

                        <small>
                            ${escaparHTML(
                                error.message
                            )}
                        </small>

                    </td>

                </tr>

            `;
        }


    } finally {

        actualizacionEnCurso =
            false;


        /*
         * SIEMPRE quitar el overlay.
         */

        ocultarCarga();
    }
}


/* ============================================================
   ESCAPAR HTML
   ============================================================ */

function escaparHTML(valor) {

    return String(
        valor ?? ""
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


/* ============================================================
   EVENTOS DE FILTROS
   ============================================================ */

if (buscarEquipo) {

    buscarEquipo.addEventListener(
        "input",
        aplicarFiltrosEquipos
    );
}


if (filtroEstadoEquipo) {

    filtroEstadoEquipo.addEventListener(
        "change",
        aplicarFiltrosEquipos
    );
}


if (filtroAreaEquipo) {

    filtroAreaEquipo.addEventListener(
        "change",
        aplicarFiltrosEquipos
    );
}


/* ============================================================
   INICIO
   ============================================================ */

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        log(
            "📄 DOM cargado."
        );


        const correcto =
            inicializarMSAL();


        if (!correcto) {

            ocultarCarga();

            return;
        }


        /*
         * Cargar dashboard.
         */

        await cargarDashboard();


        /*
         * Actualizar cada 2 minutos.
         */

        setInterval(
            cargarDashboard,
            120000
        );
    }
);


/* ============================================================
   FUNCIONES GLOBALES
   ============================================================ */

window.cargarDashboard =
    cargarDashboard;

window.cargarEquipos =
    cargarEquipos;

window.obtenerEquipos =
    obtenerEquipos;

window.obtenerMonitoreoTI =
    obtenerMonitoreoTI;

window.obtenerTokenInfraestructura =
    obtenerTokenInfraestructura;

window.cambiarPagina =
    cambiarPagina;

window.aplicarFiltrosEquipos =
    aplicarFiltrosEquipos;

window.mostrarCarga =
    mostrarCarga;

window.ocultarCarga =
    ocultarCarga;