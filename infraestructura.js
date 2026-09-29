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

    const contenedor =
        document.getElementById(
            "serviciosGrid"
        );


    if (
        !contenedor
    ) return;


    contenedor.innerHTML = "";


    if (
        !datos.length
    ) {

        contenedor.innerHTML = `

            <div class="service-card">

                <div class="service-top">

                    <div class="service-icon">
                        ⚙️
                    </div>

                    <span class="status-badge">
                        Sin datos
                    </span>

                </div>

                <h4>
                    Monitoreo TI
                </h4>

                <p>
                    No existen registros configurados en SharePoint.
                </p>

            </div>

        `;

        return;

    }


    datos.forEach(
        item => {

            const fields =
                obtenerCampos(
                    item
                );


            const servicio =
                fields.Servicio ||
                fields.Title ||
                "Servicio";


            const estado =
                normalizarEstado(
                    fields.Estado
                );


            const detalle =
                fields.Detalle ||
                "Sin detalle disponible";


            const latencia =
                fields.Latencia;


            const clase =
                clasificarEstado(
                    estado
                );


            let datoTexto =
                "Sin dato";


            if (
                latencia !== undefined &&
                latencia !== null &&
                latencia !== ""
            ) {

                datoTexto =
                    `${latencia} ms`;

            }


            contenedor.innerHTML += `

                <div class="service-card">

                    <div class="service-top">

                        <div class="service-icon">
                            ${obtenerIconoServicio(servicio)}
                        </div>

                        <span class="status-badge ${clase}">
                            ${estado}
                        </span>

                    </div>

                    <h4>
                        ${servicio}
                    </h4>

                    <p>
                        ${detalle}
                    </p>

                    <div class="service-data">

                        <span>
                            Latencia
                        </span>

                        <strong>
                            ${datoTexto}
                        </strong>

                    </div>

                </div>

            `;

        }
    );

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

function actualizarEstadoGeneral(datos) {
  const estadoGeneral = document.getElementById("estadoMonitoreo");

  if (!estadoGeneral) return;

  if (!datos || datos.length === 0) {
    estadoGeneral.textContent = "No configurado";
    estadoGeneral.className = "status-badge neutral";
    return;
  }

  const estados = datos.map(item => {
    const campos = obtenerCampos(item);
    return normalizarEstado(campos.Estado);
  });

  const tieneIncidencia = estados.some(estado =>
    ["incidencia", "error", "caido"].includes(estado)
  );

  const tieneAdvertencia = estados.some(estado =>
    ["advertencia", "warning", "pendiente"].includes(estado)
  );

  const todosNoConfigurados = estados.every(estado =>
    estado === "no configurado" || estado === ""
  );

  const todosOperativos = estados.every(estado =>
    ["operativo", "vigente", "activo", "responde"].includes(estado)
  );

  if (tieneIncidencia) {
    estadoGeneral.textContent = "Incidencia";
    estadoGeneral.className = "status-badge offline";
  } else if (tieneAdvertencia) {
    estadoGeneral.textContent = "Advertencia";
    estadoGeneral.className = "status-badge warning";
  } else if (todosNoConfigurados) {
    estadoGeneral.textContent = "No configurado";
    estadoGeneral.className = "status-badge neutral";
  } else if (todosOperativos) {
    estadoGeneral.textContent = "Operativo";
    estadoGeneral.className = "status-badge online";
  } else {
    estadoGeneral.textContent = "No configurado";
    estadoGeneral.className = "status-badge neutral";
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
   CARGAR MONITOREO
========================================= */

async function cargarMonitoreo() {

    try {

        mostrarCarga(
            "Cargando monitoreo..."
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