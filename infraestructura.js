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
   UBICACIÓN DEL EQUIPOS.JSON
========================================= */



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
        "✅ Sitio SharePoint encontrado:",
        sitio
    );


    return sitio;

}


/* =========================================
   OBTENER LISTA MONITOREOTI
========================================= */

async function obtenerMonitoreoTI() {

    try {

        console.log(
            "===================================="
        );

        console.log(
            "🚀 INICIANDO MONITOREO TI"
        );

        console.log(
            "===================================="
        );


        cambiarMensaje(
            "Conectando con Microsoft..."
        );


        /* ================================
           TOKEN
        ================================= */

        const TOKEN =
            await obtenerTokenInfraestructura();


        /* ================================
           SITIO
        ================================= */

        cambiarMensaje(
            "Conectando con SharePoint..."
        );


        const sitio =
            await obtenerSitioSharePoint(
                TOKEN
            );


        /* ================================
           LISTA
        ================================= */

        cambiarMensaje(
            "Consultando Monitoreo TI..."
        );


        const respuesta =
            await fetch(

                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${SHAREPOINT_LIST}/items?expand=fields`,

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
                "❌ Error obteniendo lista MonitoreoTI:",
                error
            );


            throw new Error(
                `MonitoreoTI: HTTP ${respuesta.status}`
            );

        }


        const data =
            await respuesta.json();


        console.log(
            "===================================="
        );

        console.log(
            "✅ DATOS DE MONITOREO TI"
        );

        console.log(
            "===================================="
        );


        console.table(
            data.value
        );


        return data.value || [];

    }

    catch (
        error
    ) {

        console.error(
            "❌ Error obteniendo Monitoreo TI:",
            error
        );


        throw error;

    }

}


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
   RENDERIZAR SERVICIOS PRINCIPALES
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


    /* =====================================
       SIN DATOS
    ===================================== */

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


    /* =====================================
       ESTADOS
    ===================================== */

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


            const card =
                document.createElement(
                    "div"
                );


            card.className =
                `service-card ${tipoEstado}`;


            card.innerHTML = `

                <div class="service-icon">
                    ${icono}
                </div>

                <div class="service-content">

                    <div class="service-status ${tipoEstado}">

                        <span class="status-dot"></span>

                        ${estado}

                    </div>

                    <h3>
                        ${servicio}
                    </h3>

                    <p>
                        ${detalle}
                    </p>

                    <div class="service-latency">

                        <span>
                            Latencia
                        </span>

                        <strong>
                            ${latenciaTexto}
                        </strong>

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
   ACTUALIZAR HORA
========================================= */

function actualizarHora() {

    const elemento =
        document.getElementById(
            "lastUpdate"
        );


    if (
        !elemento
    ) return;


    const ahora =
        new Date();


    elemento.textContent =
        ahora.toLocaleTimeString(
            "es-PE",
            {

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
   EQUIPOS
========================================= */


/* =========================================
   OBTENER EQUIPOS DESDE SHAREPOINT
========================================= */

async function obtenerEquipos() {

    try {

        console.log("🔄 Cargando inventario desde SharePoint...");

        const TOKEN = await obtenerTokenInfraestructura();

        if (!TOKEN) {
            throw new Error("No se obtuvo el token de Microsoft Graph.");
        }

        // ============================================================
        // 1. OBTENER SITIO SHAREPOINT
        // ============================================================

        const sitio = await obtenerSitioSharePoint(TOKEN);

        console.log("🏢 Sitio SharePoint:", sitio);

        if (!sitio || !sitio.id) {
            throw new Error("No se pudo obtener el sitio de SharePoint.");
        }

        // ============================================================
        // 2. OBTENER DRIVE
        // ============================================================

        const respuestaDrive = await fetch(
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

            const errorDrive = await respuestaDrive.text();

            console.error(
                "❌ Error obteniendo Drive:",
                respuestaDrive.status,
                errorDrive
            );

            throw new Error(
                `No se pudo obtener el Drive (${respuestaDrive.status})`
            );
        }

        const drive = await respuestaDrive.json();

        console.log("💾 Drive encontrado:", drive);

        if (!drive.id) {
            throw new Error("El Drive no tiene un ID válido.");
        }

        // ============================================================
        // 3. RUTA DEL EQUIPOS.JSON
        // ============================================================

        const rutaArchivo ="Procedimientos T.I/equipos.json";

        const urlArchivo =
            `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${encodeURI(rutaArchivo)}:/content`;

        console.log("📂 Ruta del archivo:", rutaArchivo);
        console.log("🌐 URL Graph:", urlArchivo);

        // ============================================================
        // 4. DESCARGAR EQUIPOS.JSON
        // ============================================================

        const respuestaArchivo = await fetch(
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

            const errorArchivo = await respuestaArchivo.text();

            console.error(
                "❌ Error obteniendo equipos.json:",
                respuestaArchivo.status,
                errorArchivo
            );

            throw new Error(
                `No se pudo obtener equipos.json (${respuestaArchivo.status})`
            );
        }

        // ============================================================
        // 5. LEER JSON
        // ============================================================

        const datos = await respuestaArchivo.json();

        console.log("📦 RESPUESTA COMPLETA DEL JSON:", datos);
        console.log("📊 ¿Es un array?:", Array.isArray(datos));
        console.log(
            "🔢 Cantidad:",
            Array.isArray(datos) ? datos.length : "NO ES ARRAY"
        );

        // ============================================================
        // 6. VALIDAR ESTRUCTURA
        // ============================================================

        if (!Array.isArray(datos)) {

            console.error(
                "❌ equipos.json no contiene directamente un array:",
                datos
            );

            throw new Error(
                "El archivo equipos.json no tiene el formato esperado."
            );
        }

        if (datos.length === 0) {

            console.warn(
                "⚠️ equipos.json fue cargado correctamente, pero está vacío."
            );

            return [];
        }

        // ============================================================
        // 7. MOSTRAR PRIMER EQUIPO PARA COMPROBAR ESTRUCTURA
        // ============================================================

        console.log(
            "🔍 Primer equipo del inventario:",
            datos[0]
        );

        console.log(
            "✅ Inventario cargado correctamente:",
            datos.length,
            "equipos"
        );

        return datos;

    } catch (error) {

        console.error(
            "❌ Error cargando inventario desde SharePoint:",
            error
        );

        const contenedor =
            document.getElementById("equiposGrid");

        if (contenedor) {

            contenedor.innerHTML = `
                <div class="error-inventario">
                    ❌ No se pudo cargar
                    <strong>el inventario desde SharePoint</strong>.
                </div>
            `;
        }

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
   CLASIFICAR ESTADO DE EQUIPO
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
   RENDERIZAR TABLA DE EQUIPOS
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


    const resultados =
        document.getElementById(
            "equiposResultados"
        );


    if (
        !tbody
    ) return;


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


        if (
            resultados
        ) {

            resultados.textContent =
                "0";

        }


        return;

    }


    if (
        empty
    ) {

        empty.style.display =
            "none";

    }


    if (
        resultados
    ) {

        resultados.textContent =
            equipos.length;

    }


    equipos.forEach(
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

}


/* =========================================
   ACTUALIZAR RESUMEN DE EQUIPOS
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
   CREAR FILTRO DE ÁREAS
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
        buscar
    ) {

        buscar.addEventListener(
            "input",
            aplicarFiltrosEquipos
        );

    }


    if (
        filtroEstado
    ) {

        filtroEstado.addEventListener(
            "change",
            aplicarFiltrosEquipos
        );

    }


    if (
        filtroArea
    ) {

        filtroArea.addEventListener(
            "change",
            aplicarFiltrosEquipos
        );

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
           EQUIPOS
        ================================= */

        cambiarMensaje(
            "Cargando inventario de equipos..."
        );


        await cargarEquipos();


        /* =================================
           HORA
        ================================= */

        actualizarHora();


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
   CADA 5 MINUTOS
========================================= */

setInterval(
    function () {

        cargarMonitoreo();

    },
    300000
);