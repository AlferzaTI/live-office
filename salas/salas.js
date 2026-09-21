import { obtenerToken } from "./salas-api.js";

/* =========================================================
   CONFIGURACIÓN
========================================================= */

const SITE_PATH =
    "https://graph.microsoft.com/v1.0/sites/alferzaholding-my.sharepoint.com:/personal/soporte1_alferza_pe";

const NOMBRE_LISTA = "ReservaSalas";

let reservas = [];


/* =========================================================
   ELEMENTOS HTML
========================================================= */

const elementos = {
    totalSalas: document.getElementById("totalSalas"),
    reservasHoy: document.getElementById("reservasHoy"),
    reservasActivas: document.getElementById("reservasActivas"),
    reservasCanceladas: document.getElementById("reservasCanceladas"),

    sala2Reservas: document.getElementById("sala2Reservas"),
    sala3Reservas: document.getElementById("sala3Reservas"),
    comedorReservas: document.getElementById("comedorReservas"),

    fechaActual: document.getElementById("fechaActual"),

    tablaReservas: document.getElementById("tablaReservas"),
    emptyState: document.getElementById("emptyState"),
    totalReservasLabel: document.getElementById("totalReservasLabel"),

    refreshButton: document.getElementById("refreshSalas")
};


/* =========================================================
   FECHAS
========================================================= */

function obtenerFechaActualISO() {

    const ahora = new Date();

    const year = ahora.getFullYear();

    const month = String(
        ahora.getMonth() + 1
    ).padStart(2, "0");

    const day = String(
        ahora.getDate()
    ).padStart(2, "0");

    return `${year}-${month}-${day}`;
}


function obtenerFechaReservaISO(reserva) {

    if (!reserva.FechaReserva) {
        return "";
    }

    /*
        SharePoint puede devolver:

        2026-09-21T07:00:00Z
        2026-09-21T00:00:00Z
        2026-09-21

        Nosotros solamente necesitamos
        la parte YYYY-MM-DD.
    */

    return String(
        reserva.FechaReserva
    ).substring(0, 10);
}


function formatearFecha(fecha) {

    if (!fecha) {
        return "-";
    }

    const fechaISO =
        obtenerFechaReservaISO({
            FechaReserva: fecha
        });

    const partes =
        fechaISO.split("-");

    if (partes.length !== 3) {
        return fechaISO;
    }

    return `${partes[2]}/${partes[1]}/${partes[0]}`;
}


/* =========================================================
   HORAS
========================================================= */

function convertirHoraAMinutos(hora) {

    if (!hora) {
        return null;
    }

    const texto =
        String(hora).trim();

    /*
        Acepta:

        08:00
        08:30
        14:00
        08:00:00
        08:00:00.0000000
    */

    const match =
        texto.match(/^(\d{1,2}):(\d{2})/);

    if (!match) {
        return null;
    }

    const horas =
        Number(match[1]);

    const minutos =
        Number(match[2]);

    if (
        Number.isNaN(horas) ||
        Number.isNaN(minutos)
    ) {
        return null;
    }

    return horas * 60 + minutos;
}


function formatearHora(hora) {

    if (!hora) {
        return "--:--";
    }

    const texto =
        String(hora).trim();

    const match =
        texto.match(/^(\d{1,2}):(\d{2})/);

    if (!match) {
        return texto;
    }

    return `${String(match[1]).padStart(2, "0")}:${match[2]}`;
}


/* =========================================================
   NORMALIZAR DATOS DE SHAREPOINT
========================================================= */

function normalizarReservas(items) {

    return items.map(item => {

        const fields =
            item.fields || {};

        return {

            id: item.id,

            Sala:
                fields.Sala ?? "",

            FechaReserva:
                fields.FechaReserva ?? "",

            HoraInicio:
                fields.HoraInicio ?? "",

            HoraFin:
                fields.HoraFin ?? "",

            Motivo:
                fields.Motivo ?? "",

            Solicitante:
                fields.Solicitante ?? "",

            CorreoSolicitante:
                fields.CorreoSolicitante ?? "",

            Estado:
                fields.Estado ?? "",

            FechaCreacion:
                fields.FechaCreacion ?? "",

            CanceladoPor:
                fields.CanceladoPor ?? "",

            FechaCancelacion:
                fields.FechaCancelacion ?? "",

            IDReserva:
                fields.IDReserva ?? "",

            BloqueInicio:
                fields.BloqueInicio ?? "",

            BloqueFin:
                fields.BloqueFin ?? ""
        };
    });
}


/* =========================================================
   OBTENER SITIO
========================================================= */

async function obtenerSitio(token) {

    const respuesta =
        await fetch(
            SITE_PATH,
            {
                headers: {
                    Authorization:
                        `Bearer ${token}`
                }
            }
        );

    if (!respuesta.ok) {

        const texto =
            await respuesta.text();

        throw new Error(
            `Error obteniendo sitio: ${respuesta.status} ${texto}`
        );
    }

    return respuesta.json();
}


/* =========================================================
   OBTENER LISTA
========================================================= */

async function obtenerLista(
    token,
    siteId
) {

    const url =
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${encodeURIComponent(NOMBRE_LISTA)}?expand=columns`;

    const respuesta =
        await fetch(
            url,
            {
                headers: {
                    Authorization:
                        `Bearer ${token}`
                }
            }
        );

    if (!respuesta.ok) {

        const texto =
            await respuesta.text();

        throw new Error(
            `Error obteniendo lista: ${respuesta.status} ${texto}`
        );
    }

    return respuesta.json();
}


/* =========================================================
   OBTENER ITEMS
========================================================= */

async function obtenerItems(
    token,
    siteId,
    listId
) {

    let url =
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`;

    const todosLosItems = [];

    while (url) {

        const respuesta =
            await fetch(
                url,
                {
                    headers: {
                        Authorization:
                            `Bearer ${token}`
                    }
                }
            );

        if (!respuesta.ok) {

            const texto =
                await respuesta.text();

            throw new Error(
                `Error obteniendo reservas: ${respuesta.status} ${texto}`
            );
        }

        const datos =
            await respuesta.json();

        if (Array.isArray(datos.value)) {

            todosLosItems.push(
                ...datos.value
            );
        }

        url =
            datos["@odata.nextLink"] || null;
    }

    return todosLosItems;
}


/* =========================================================
   OBTENER TODAS LAS RESERVAS
========================================================= */

async function obtenerReservas() {

    const token =
        await obtenerToken();

    const sitio =
        await obtenerSitio(token);

    console.log(
        "Sitio obtenido:",
        sitio
    );

    const lista =
        await obtenerLista(
            token,
            sitio.id
        );

    console.log(
        "Lista obtenida:",
        lista
    );

    const items =
        await obtenerItems(
            token,
            sitio.id,
            lista.id
        );

    console.log(
        "Items obtenidos:",
        items.length
    );

    const datos =
        normalizarReservas(items);

    console.log(
        "Reservas normalizadas:",
        datos
    );

    return datos;
}


/* =========================================================
   RESERVAS DE HOY
========================================================= */

function obtenerReservasHoy() {

    const hoy =
        obtenerFechaActualISO();

    return reservas.filter(
        reserva =>
            obtenerFechaReservaISO(reserva) === hoy
    );
}


/* =========================================================
   RESERVAS ACTIVAS DE HOY
========================================================= */

function obtenerReservasActivasHoy() {

    return obtenerReservasHoy()
        .filter(
            reserva =>
                String(reserva.Estado)
                    .trim()
                    .toLowerCase() === "activa"
        );
}


/* =========================================================
   ACTUALIZAR FECHA DEL HEADER
========================================================= */

function actualizarFecha() {

    if (!elementos.fechaActual) {
        return;
    }

    const ahora =
        new Date();

    const opciones = {
        day: "2-digit",
        month: "short",
        year: "numeric"
    };

    let texto =
        ahora.toLocaleDateString(
            "es-PE",
            opciones
        );

    texto =
        texto
            .replace(".", "")
            .toUpperCase();

    elementos.fechaActual.textContent =
        texto;
}


/* =========================================================
   ACTUALIZAR CONTADORES
========================================================= */

function actualizarContadores() {

    const reservasHoy =
        obtenerReservasHoy();

    const activas =
        reservasHoy.filter(
            reserva =>
                String(reserva.Estado)
                    .trim()
                    .toLowerCase() === "activa"
        );

    const canceladas =
        reservasHoy.filter(
            reserva =>
                String(reserva.Estado)
                    .trim()
                    .toLowerCase() === "cancelada"
        );

    if (elementos.totalSalas) {

        elementos.totalSalas.textContent =
            "3";
    }

    if (elementos.reservasHoy) {

        elementos.reservasHoy.textContent =
            reservasHoy.length;
    }

    if (elementos.reservasActivas) {

        elementos.reservasActivas.textContent =
            activas.length;
    }

    if (elementos.reservasCanceladas) {

        elementos.reservasCanceladas.textContent =
            canceladas.length;
    }

    console.log(
        "Reservas de hoy:",
        reservasHoy.length
    );

    console.log(
        "Activas:",
        activas.length
    );

    console.log(
        "Canceladas:",
        canceladas.length
    );
}


/* =========================================================
   CONTADORES POR SALA
========================================================= */

function actualizarContadoresSalas() {

    const reservasHoy =
        obtenerReservasHoy();

    const contar =
        nombreSala =>
            reservasHoy.filter(
                reserva =>
                    String(reserva.Sala)
                        .trim()
                        .toLowerCase() ===
                    nombreSala.toLowerCase()
            ).length;

    if (elementos.sala2Reservas) {

        elementos.sala2Reservas.textContent =
            contar("Sala 2");
    }

    if (elementos.sala3Reservas) {

        elementos.sala3Reservas.textContent =
            contar("Sala 3");
    }

    if (elementos.comedorReservas) {

        elementos.comedorReservas.textContent =
            contar("Comedor");
    }
}


/* =========================================================
   ESTADO ACTUAL DE UNA SALA
========================================================= */

function obtenerEstadoSala(nombreSala) {

    const ahora =
        new Date();

    const horaActual =
        ahora.getHours() * 60 +
        ahora.getMinutes();

    const reservasSala =
        obtenerReservasActivasHoy()
            .filter(
                reserva =>
                    String(reserva.Sala)
                        .trim()
                        .toLowerCase() ===
                    nombreSala.toLowerCase()
            );

    const reservaActual =
        reservasSala.find(
            reserva => {

                const inicio =
                    convertirHoraAMinutos(
                        reserva.HoraInicio
                    );

                const fin =
                    convertirHoraAMinutos(
                        reserva.HoraFin
                    );

                if (
                    inicio === null ||
                    fin === null
                ) {
                    return false;
                }

                return (
                    horaActual >= inicio &&
                    horaActual < fin
                );
            }
        );

    if (reservaActual) {

        return {
            estado: "Ocupada",
            reserva: reservaActual
        };
    }

    return {
        estado: "Disponible",
        reserva: null
    };
}


/* =========================================================
   ACTUALIZAR ESTADOS VISUALES DE LAS SALAS
========================================================= */

function actualizarEstadosVisualesSalas() {

    const tarjetas =
        document.querySelectorAll(
            ".sala-card"
        );

    tarjetas.forEach(
        tarjeta => {

            const titulo =
                tarjeta.querySelector("h3");

            if (!titulo) {
                return;
            }

            const nombreSala =
                titulo.textContent.trim();

            const estadoTop =
                tarjeta.querySelector(
                    ".sala-status"
                );

            const estadoFooter =
                tarjeta.querySelector(
                    ".sala-footer strong"
                );

            if (
                !estadoTop ||
                !estadoFooter
            ) {
                return;
            }

            const resultado =
                obtenerEstadoSala(
                    nombreSala
                );

            if (
                resultado.estado === "Ocupada"
            ) {

                estadoTop.classList.remove(
                    "available"
                );

                estadoTop.classList.add(
                    "busy"
                );

                estadoTop.innerHTML =
                    "<i></i> Ocupada";

                estadoFooter.classList.remove(
                    "text-green"
                );

                estadoFooter.classList.add(
                    "text-red"
                );

                estadoFooter.textContent =
                    "Ocupada";

            } else {

                estadoTop.classList.remove(
                    "busy"
                );

                estadoTop.classList.add(
                    "available"
                );

                estadoTop.innerHTML =
                    "<i></i> Disponible";

                estadoFooter.classList.remove(
                    "text-red"
                );

                estadoFooter.classList.add(
                    "text-green"
                );

                estadoFooter.textContent =
                    "Disponible";
            }
        }
    );
}


/* =========================================================
   ESCAPAR HTML
========================================================= */

function escaparHTML(valor) {

    return String(valor ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}


/* =========================================================
   RENDERIZAR TABLA
========================================================= */

function renderizarTabla() {

    if (!elementos.tablaReservas) {
        return;
    }

    const reservasHoy =
        obtenerReservasHoy();

    elementos.tablaReservas.innerHTML =
        "";

    if (elementos.totalReservasLabel) {

        elementos.totalReservasLabel.textContent =
            reservasHoy.length;
    }

    /*
        Ordenar por hora de inicio.
    */

    const ordenadas =
        [...reservasHoy].sort(
            (a, b) => {

                const horaA =
                    convertirHoraAMinutos(
                        a.HoraInicio
                    ) ?? 9999;

                const horaB =
                    convertirHoraAMinutos(
                        b.HoraInicio
                    ) ?? 9999;

                return horaA - horaB;
            }
        );

    if (ordenadas.length === 0) {

        if (elementos.emptyState) {
            elementos.emptyState.hidden =
                false;
        }

        return;
    }

    if (elementos.emptyState) {
        elementos.emptyState.hidden =
            true;
    }


    ordenadas.forEach(
        reserva => {

            const fila =
                document.createElement("tr");

            const estadoNormalizado =
                String(
                    reserva.Estado
                )
                    .trim()
                    .toLowerCase();

            const esActiva =
                estadoNormalizado ===
                "activa";

            const claseEstado =
                esActiva
                    ? "estado-activa"
                    : "estado-cancelada";

            const textoEstado =
                esActiva
                    ? "Activa"
                    : "Cancelada";


            fila.innerHTML = `

                <td>
                    <div class="sala-cell">
                        ${escaparHTML(reserva.Sala)}
                    </div>
                </td>

                <td>
                    ${escaparHTML(
                        formatearFecha(
                            reserva.FechaReserva
                        )
                    )}
                </td>

                <td>
                    ${escaparHTML(
                        formatearHora(
                            reserva.HoraInicio
                        )
                    )}
                    -
                    ${escaparHTML(
                        formatearHora(
                            reserva.HoraFin
                        )
                    )}
                </td>

                <td>

                    <div class="solicitante-cell">

                        <div class="solicitante-avatar">
                            ${escaparHTML(
                                obtenerIniciales(
                                    reserva.Solicitante
                                )
                            )}
                        </div>

                        <div>

                            <div class="solicitante-nombre">
                                ${escaparHTML(
                                    reserva.Solicitante ||
                                    "Sin solicitante"
                                )}
                            </div>

                            <div class="solicitante-motivo">
                                ${escaparHTML(
                                    reserva.Motivo ||
                                    "Sin motivo"
                                )}
                            </div>

                        </div>

                    </div>

                </td>

                <td>

                    <span class="estado ${claseEstado}">

                        <i class="estado-dot"></i>

                        ${textoEstado}

                    </span>

                </td>

            `;

            elementos.tablaReservas
                .appendChild(fila);
        }
    );
}


/* =========================================================
   INICIALES
========================================================= */

function obtenerIniciales(nombre) {

    if (!nombre) {
        return "—";
    }

    const palabras =
        String(nombre)
            .trim()
            .split(/\s+/)
            .filter(Boolean);

    if (palabras.length === 1) {

        return palabras[0]
            .substring(0, 2)
            .toUpperCase();
    }

    return (
        palabras[0][0] +
        palabras[1][0]
    ).toUpperCase();
}


/* =========================================================
   CARGAR TODO
========================================================= */

async function cargarReservas() {

    try {

        console.log(
            "Cargando reservas..."
        );

        if (elementos.tablaReservas) {

            elementos.tablaReservas.innerHTML = `
                <tr>
                    <td colspan="5" style="text-align:center;">
                        Cargando reservas...
                    </td>
                </tr>
            `;
        }

        reservas =
            await obtenerReservas();

        console.log(
            "TOTAL RESERVAS:",
            reservas.length
        );

        console.log(
            "FECHA ACTUAL:",
            obtenerFechaActualISO()
        );

        console.log(
            "FECHAS EN DATOS:",
            reservas.map(
                reserva =>
                    reserva.FechaReserva
            )
        );

        actualizarFecha();

        actualizarContadores();

        actualizarContadoresSalas();

        renderizarTabla();

        actualizarEstadosVisualesSalas();

        console.log(
            "Reservas cargadas correctamente."
        );

    } catch (error) {

        console.error(
            "ERROR CARGANDO RESERVAS:",
            error
        );

        if (elementos.tablaReservas) {

            elementos.tablaReservas.innerHTML = `
                <tr>
                    <td colspan="5" style="text-align:center;">
                        Error al cargar las reservas.
                    </td>
                </tr>
            `;
        }

        if (elementos.emptyState) {
            elementos.emptyState.hidden =
                true;
        }
    }
}


/* =========================================================
   REFRESCAR
========================================================= */

async function refrescarTodo() {

    await cargarReservas();
}


/* =========================================================
   INICIO
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        await cargarReservas();

        /*
            Actualiza automáticamente cada minuto
            para que el estado de las salas cambie
            de Disponible <-> Ocupada.
        */

        setInterval(
            async () => {

                await cargarReservas();

            },
            60 * 1000
        );

        /*
            Si posteriormente agregas un botón
            con id="refreshSalas", funcionará
            automáticamente.
        */

        if (elementos.refreshButton) {

            elementos.refreshButton.addEventListener(
                "click",
                refrescarTodo
            );
        }
    }
);