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


const msalInstanceInfraestructura =
    new msal.PublicClientApplication(
        msalConfigInfraestructura
    );


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

async function iniciarSesion() {

    try {

        const cuentas =
            msalInstanceInfraestructura
                .getAllAccounts();


        if (
            cuentas.length > 0
        ) {

            cuentaActual =
                cuentas[0];

        } else {

            const loginResponse =
                await msalInstanceInfraestructura
                    .loginPopup({

                        scopes: [
                            "User.Read",
                            "Sites.Read.All",
                            "Sites.ReadWrite.All"
                        ]

                    });

            cuentaActual =
                loginResponse.account;
        }


        const tokenResponse =
            await msalInstanceInfraestructura
                .acquireTokenSilent({

                    scopes: [
                        "User.Read",
                        "Sites.Read.All",
                        "Sites.ReadWrite.All"
                    ],

                    account:
                        cuentaActual
                });


        tokenActual =
            tokenResponse.accessToken;


        return tokenActual;


    } catch (error) {

        console.error(
            "Error de autenticación:",
            error
        );


        try {

            const tokenPopup =
                await msalInstanceInfraestructura
                    .acquireTokenPopup({

                        scopes: [
                            "User.Read",
                            "Sites.Read.All",
                            "Sites.ReadWrite.All"
                        ]

                    });


            tokenActual =
                tokenPopup.accessToken;

            cuentaActual =
                tokenPopup.account;


            return tokenActual;


        } catch (popupError) {

            console.error(
                "Error obteniendo token:",
                popupError
            );

            throw popupError;
        }
    }
}


/* =========================================
   OBTENER SITIO SHAREPOINT
========================================= */

async function obtenerSitioSharePoint(
    token
) {

    if (
        sitioSharePoint
    ) {
        return sitioSharePoint;
    }


    const url =
        `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOSTNAME}:${SHAREPOINT_SITE}`;


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
            `No se pudo obtener el sitio SharePoint. ${response.status} ${texto}`
        );
    }


    sitioSharePoint =
        await response.json();


    return sitioSharePoint;
}


/* =========================================
   OBTENER DRIVE
========================================= */

async function obtenerDriveSharePoint(
    token
) {

    if (
        driveSharePoint
    ) {
        return driveSharePoint;
    }


    const sitio =
        await obtenerSitioSharePoint(
            token
        );


    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/drive`;


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
            `No se pudo obtener el drive de SharePoint. ${response.status} ${texto}`
        );
    }


    driveSharePoint =
        await response.json();


    return driveSharePoint;
}


/* =========================================
   OBTENER ARCHIVO EQUIPOS
========================================= */

async function obtenerEquiposDesdeSharePoint(
    token
) {

    const drive =
        await obtenerDriveSharePoint(
            token
        );


    const rutaCodificada =
        RUTA_EQUIPOS
            .split("/")
            .map(
                parte =>
                    encodeURIComponent(
                        parte
                    )
            )
            .join("/");


    const url =
        `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${rutaCodificada}:/content`;


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
            `No se pudo descargar equipos.json. ${response.status} ${texto}`
        );
    }


    const datos =
        await response.json();


    if (
        !Array.isArray(datos)
    ) {

        throw new Error(
            "El archivo equipos.json no contiene un arreglo válido."
        );
    }


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

async function obtenerTrabajadores(
    token
) {

    /*
     * Si ya fueron cargados,
     * no hacemos otra petición.
     */

    if (
        trabajadoresCargados
    ) {

        return trabajadoresPorEquipo;
    }


    /*
     * Si otra ejecución ya está cargando,
     * esperamos la misma promesa.
     */

    if (
        cargandoTrabajadores
    ) {

        return trabajadoresPorEquipo;
    }


    cargandoTrabajadores =
        true;


    try {

        const sitio =
            await obtenerSitioSharePoint(
                token
            );


        const lista =
            await obtenerListaTrabajadores(
                token
            );


        /*
         * IMPORTANTE:
         *
         * Solo solicitamos los campos
         * necesarios.
         *
         * Ya no usamos:
         *
         * expand=fields
         *
         * para traer toda la información.
         */

        const url =
            `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items?expand=fields($select=Equipo1,Trabajador)&$top=999`;


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
                `No se pudieron obtener las asignaciones de trabajadores. ${response.status} ${texto}`
            );
        }


        const data =
            await response.json();


        trabajadoresData =
            data.value || [];


        trabajadoresPorEquipo =
            {};


        /*
         * Como solicitamos los nombres visibles
         * de las columnas, primero intentamos
         * acceder directamente.
         */

        trabajadoresData.forEach(
            item => {

                const fields =
                    item.fields || {};


                const equipo =
                    fields.Equipo1;


                const trabajador =
                    fields.Trabajador;


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


        /*
         * Si SharePoint devolvió los campos
         * con nombres internos diferentes,
         * usamos el método anterior como respaldo.
         */

        if (
            trabajadoresData.length > 0 &&
            Object.keys(
                trabajadoresPorEquipo
            ).length === 0
        ) {

            await cargarTrabajadoresConColumnas(
                token,
                sitio,
                lista
            );
        }


        trabajadoresCargados =
            true;


        console.log(
            `Asignaciones de trabajadores cargadas: ${Object.keys(trabajadoresPorEquipo).length}`
        );


        /*
         * Actualizamos la tabla únicamente
         * después de terminar la consulta.
         *
         * La página ya estaba visible.
         */

        if (
            equiposData.length > 0
        ) {

            renderizarTablaEquipos();
        }


        return trabajadoresPorEquipo;


    } catch (error) {

        console.warn(
            "No se pudieron cargar las asignaciones de trabajadores:",
            error
        );


        /*
         * IMPORTANTE:
         *
         * No lanzamos el error.
         *
         * Si falla esta parte,
         * el inventario sigue funcionando.
         */

        return trabajadoresPorEquipo;


    } finally {

        cargandoTrabajadores =
            false;
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

async function obtenerMonitoreoTI(
    token
) {

    const sitio =
        await obtenerSitioSharePoint(
            token
        );


    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${encodeURIComponent(SHAREPOINT_LIST_SERVICIOS)}/items?expand=fields&$top=999`;


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
            `No se pudo obtener MonitoreoTI. ${response.status} ${texto}`
        );
    }


    const data =
        await response.json();


    return data.value || [];
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


                    const revision =
                        fields.UltimaRevision ||
                        fields["ÚltimaRevision"] ||
                        fields.Fecha ||
                        "--";


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

function aplicarFiltrosEquipos() {

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


    paginaActual =
        1;


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


    modalEquipoNombre.textContent =
        nombreEquipo;


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

    const nombre =
        trabajadorNombre.value.trim();


    if (
        !equipoSeleccionado
    ) {

        mostrarErrorTrabajador(
            "No se seleccionó ningún equipo."
        );

        return;
    }


    if (
        !nombre
    ) {

        mostrarErrorTrabajador(
            "Ingrese el nombre del trabajador."
        );


        trabajadorNombre.focus();


        return;
    }


    try {

        guardarTrabajador.disabled =
            true;


        trabajadorError.style.display =
            "none";


        guardarTrabajador.textContent =
            "Guardando...";


        const sitio =
            await obtenerSitioSharePoint(
                tokenActual
            );


        const lista =
            await obtenerListaTrabajadores(
                tokenActual
            );


        const columnas =
            await obtenerColumnasTrabajadores(
                tokenActual
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
            !campoEquipo
        ) {

            throw new Error(
                'No se encontró la columna "Equipo1".'
            );
        }


        if (
            !campoTrabajador
        ) {

            throw new Error(
                'No se encontró la columna "Trabajador".'
            );
        }


        /*
         * Verificamos si ya existe una asignación.
         */

        const asignacionActual =
            obtenerTrabajadorEquipo(
                equipoSeleccionado
            );


        /*
         * Si no estaba en memoria,
         * hacemos una consulta antes de crear
         * para evitar duplicados.
         */

        let asignacionFinal =
            asignacionActual;


        if (
            !asignacionFinal &&
            !trabajadoresCargados
        ) {

            await obtenerTrabajadores(
                tokenActual
            );


            asignacionFinal =
                obtenerTrabajadorEquipo(
                    equipoSeleccionado
                );
        }


        const fields = {};


        fields[campoEquipo] =
            equipoSeleccionado;


        fields[campoTrabajador] =
            nombre;


        /*
         * Title solo si es obligatorio.
         */

        const columnaTitle =
            columnas.find(
                columna =>
                    normalizarTexto(
                        columna.displayName
                    ) === "title" ||
                    columna.name === "Title"
            );


        if (
            columnaTitle &&
            columnaTitle.required === true
        ) {

            fields[
                columnaTitle.name
            ] =
                equipoSeleccionado;
        }


        let response;


        /* =====================================
           ACTUALIZAR
        ====================================== */

        if (
            asignacionFinal &&
            asignacionFinal.itemId
        ) {

            const url =
                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items/${asignacionFinal.itemId}/fields`;


            response =
                await fetch(
                    url,
                    {

                        method:
                            "PATCH",

                        headers: {

                            Authorization:
                                `Bearer ${tokenActual}`,

                            "Content-Type":
                                "application/json"
                        },

                        body:
                            JSON.stringify(
                                fields
                            )
                    }
                );


        } else {

            /* =====================================
               CREAR
            ====================================== */

            const url =
                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items`;


            response =
                await fetch(
                    url,
                    {

                        method:
                            "POST",

                        headers: {

                            Authorization:
                                `Bearer ${tokenActual}`,

                            "Content-Type":
                                "application/json"
                        },

                        body:
                            JSON.stringify({
                                fields:
                                    fields
                            })
                    }
                );
        }


        if (
            !response.ok
        ) {

            const texto =
                await response.text();


            throw new Error(
                `SharePoint respondió ${response.status}: ${texto}`
            );
        }


        const resultado =
            await response.json();


        /*
         * Actualizar memoria local.
         */

        trabajadoresPorEquipo[
            normalizarTexto(
                equipoSeleccionado
            )
        ] = {

            nombre:
                nombre,

            itemId:
                resultado.id ||
                asignacionFinal?.itemId,

            equipo:
                equipoSeleccionado
        };


        trabajadoresCargados =
            true;


        cerrarModalTrabajador();


        renderizarTablaEquipos();


    } catch (error) {

        console.error(
            "Error guardando trabajador:",
            error
        );


        mostrarErrorTrabajador(
            obtenerMensajeError(
                error
            )
        );


    } finally {

        guardarTrabajador.disabled =
            false;


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

    buscarEquipo.addEventListener(
        "input",
        () => {

            aplicarFiltrosEquipos();
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

async function cargarEquipos() {

    try {

        equiposEstado.textContent =
            "● ACTUALIZANDO";


        equiposEstado.className =
            "section-status warning";


        /*
         * SOLO OBTENEMOS equipos.json.
         */

        const equipos =
            await obtenerEquiposDesdeSharePoint(
                tokenActual
            );


        if (
            !Array.isArray(equipos)
        ) {

            throw new Error(
                "equipos.json no contiene un arreglo válido."
            );
        }


        equiposData =
            equipos;


        equiposFiltrados =
            [...equiposData];


        paginaActual =
            1;


        /*
         * Actualizar resumen.
         */

        actualizarResumenEquipos();


        actualizarFiltroAreas();


        /*
         * PINTAR INMEDIATAMENTE.
         */

        renderizarTablaEquipos();


        /*
         * Fecha.
         */

        lastUpdate.textContent =
            obtenerFechaActual();


        equiposEstado.textContent =
            "● ACTUALIZADO";


        equiposEstado.className =
            "section-status online";


        console.log(
            `Inventario cargado: ${equiposData.length} equipos`
        );


    } catch (error) {

        console.error(
            "Error cargando equipos:",
            error
        );


        equiposEstado.textContent =
            "● ERROR";


        equiposEstado.className =
            "section-status offline";


        equiposTableBody.innerHTML = `
            <tr>
                <td
                    colspan="7"
                    class="equipos-loading"
                >

                    <div class="table-loading">

                        <span>
                            No se pudo cargar el inventario.
                        </span>

                    </div>

                </td>
            </tr>
        `;


        if (
            equiposResultados
        ) {

            equiposResultados.textContent =
                "Error al obtener los equipos.";
        }
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

async function cargarDashboard() {

    try {

        mostrarLoading(
            "Iniciando sesión..."
        );


        await iniciarSesion();


        mostrarLoading(
            "Conectando con SharePoint..."
        );


        await obtenerSitioSharePoint(
            tokenActual
        );


        /*
         * =====================================
         * SERVICIOS + INVENTARIO
         * =====================================
         *
         * Se ejecutan en paralelo.
         *
         * Antes:
         *
         * servicios → esperar → equipos
         *
         * Ahora:
         *
         * servicios ┐
         *           ├── simultáneamente
         * equipos   ┘
         *
         * Esto reduce el tiempo total.
         */

        mostrarLoading(
            "Cargando monitoreo..."
        );


        await Promise.all([
            cargarServicios(),
            cargarEquipos()
        ]);


        /*
         * Ocultamos el loading apenas
         * servicios + equipos están listos.
         */

        ocultarLoading();


        /*
         * =====================================
         * TRABAJADORES EN SEGUNDO PLANO
         * =====================================
         *
         * IMPORTANTE:
         *
         * NO usamos await.
         *
         * La página ya está funcionando.
         *
         * Si SharePoint tarda:
         * NO importa.
         *
         * Si falla:
         * NO afecta el inventario.
         */

        obtenerTrabajadores(
            tokenActual
        ).catch(
            error => {

                console.warn(
                    "Carga secundaria de trabajadores falló:",
                    error
                );
            }
        );


    } catch (error) {

        console.error(
            "Error inicializando infraestructura:",
            error
        );


        if (
            loadingMessage
        ) {

            loadingMessage.textContent =
                "No se pudo cargar el monitoreo.";
        }


        setTimeout(
            () => {

                ocultarLoading();

            },
            2500
        );
    }
}


/* =========================================
   ACTUALIZACIÓN AUTOMÁTICA
========================================= */

async function actualizarDashboard() {

    try {

        const cuentas =
            msalInstanceInfraestructura
                .getAllAccounts();


        if (
            !cuentas.length
        ) {

            return;
        }


        cuentaActual =
            cuentas[0];


        const tokenResponse =
            await msalInstanceInfraestructura
                .acquireTokenSilent({

                    scopes: [
                        "User.Read",
                        "Sites.Read.All",
                        "Sites.ReadWrite.All"
                    ],

                    account:
                        cuentaActual
                });


        tokenActual =
            tokenResponse.accessToken;


        /*
         * Solo actualizamos:
         *
         * - Servicios
         * - equipos.json
         *
         * NO volvemos a descargar
         * TrabajadoresEquipos.
         */

        await Promise.all([
            cargarServicios(),
            cargarEquipos()
        ]);


    } catch (error) {

        console.error(
            "Error en actualización automática:",
            error
        );
    }
}


/* =========================================
   INICIO
========================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        cargarDashboard();


        setInterval(
            actualizarDashboard,
            INTERVALO_ACTUALIZACION
        );
    }
);