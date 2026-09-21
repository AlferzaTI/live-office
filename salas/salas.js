import { obtenerToken } from "./salas-api.js";

/* =========================================================
   ALFERZA LIVE OFFICE
   MICROSOFT LISTS - RESERVA DE SALAS
========================================================= */


/* =========================================================
   CONFIGURACIÓN
========================================================= */

const SITE_PATH =
    "https://graph.microsoft.com/v1.0/sites/alferzaholding-my.sharepoint.com:/personal/soporte1_alferza_pe";

const NOMBRE_LISTA = "ReservaSalas";


/* =========================================================
   ESTADO DE LA APLICACIÓN
========================================================= */

let reservas = [];


/* =========================================================
   ELEMENTOS DEL DOM
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
   FECHA ACTUAL
========================================================= */

function obtenerFechaActualISO() {

    const ahora = new Date();

    const year =
        ahora.getFullYear();

    const month =
        String(
            ahora.getMonth() + 1
        ).padStart(2, "0");

    const day =
        String(
            ahora.getDate()
        ).padStart(2, "0");

    return `${year}-${month}-${day}`;
}


/* =========================================================
   FORMATEAR FECHA
========================================================= */

function formatearFecha(fecha) {

    if (!fecha) {
        return "-";
    }

    const fechaISO =
        String(fecha).substring(0, 10);

    const partes =
        fechaISO.split("-");

    if (partes.length !== 3) {
        return fecha;
    }

    return `${partes[2]}/${partes[1]}/${partes[0]}`;
}


/* =========================================================
   OBTENER FECHA DE RESERVA EN ISO
========================================================= */

function obtenerFechaReservaISO(reserva) {

    if (!reserva.FechaReserva) {
        return "";
    }

    return String(
        reserva.FechaReserva
    ).substring(0, 10);
}


/* =========================================================
   ACTUALIZAR FECHA MOSTRADA
========================================================= */

function actualizarFecha() {

    if (!elementos.fechaActual) {
        return;
    }

    const ahora =
        new Date();

    const opciones = {
        weekday: "long",
        day: "numeric",
        month: "long",
        year: "numeric"
    };

    const fecha =
        ahora.toLocaleDateString(
            "es-PE",
            opciones
        );

    elementos.fechaActual.textContent =
        fecha.charAt(0).toUpperCase() +
        fecha.slice(1);
}


/* =========================================================
   OBTENER SITE
========================================================= */

async function obtenerSite(token) {

    const respuesta =
        await fetch(
            SITE_PATH,
            {
                method: "GET",

                headers: {
                    Authorization: `Bearer ${token}`,
                    Accept: "application/json"
                }
            }
        );


    if (!respuesta.ok) {

        const error =
            await respuesta.text();

        throw new Error(
            `Error obteniendo SharePoint Site (${respuesta.status}): ${error}`
        );

    }


    return await respuesta.json();
}


/* =========================================================
   OBTENER LISTA
========================================================= */

async function obtenerLista(
    token,
    siteId
) {

    const url =
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${NOMBRE_LISTA}?expand=columns`;


    const respuesta =
        await fetch(
            url,
            {
                method: "GET",

                headers: {
                    Authorization: `Bearer ${token}`,
                    Accept: "application/json"
                }
            }
        );


    if (!respuesta.ok) {

        const error =
            await respuesta.text();

        throw new Error(
            `Error obteniendo la lista (${respuesta.status}): ${error}`
        );

    }


    return await respuesta.json();
}


/* =========================================================
   OBTENER REGISTROS
========================================================= */

async function obtenerRegistros(
    token,
    siteId,
    listaId
) {

    const url =
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items?expand=fields`;


    const respuesta =
        await fetch(
            url,
            {
                method: "GET",

                headers: {
                    Authorization: `Bearer ${token}`,
                    Accept: "application/json"
                }
            }
        );


    if (!respuesta.ok) {

        const error =
            await respuesta.text();

        throw new Error(
            `Error obteniendo reservas (${respuesta.status}): ${error}`
        );

    }


    const datos =
        await respuesta.json();


    let items =
        datos.value || [];


    /*
    ======================================
    PAGINACIÓN
    ======================================
    */

    let siguiente =
        datos["@odata.nextLink"];


    while (siguiente) {

        const respuestaSiguiente =
            await fetch(
                siguiente,
                {
                    method: "GET",

                    headers: {
                        Authorization: `Bearer ${token}`,
                        Accept: "application/json"
                    }
                }
            );


        if (!respuestaSiguiente.ok) {

            const error =
                await respuestaSiguiente.text();

            throw new Error(
                `Error obteniendo página adicional (${respuestaSiguiente.status}): ${error}`
            );

        }


        const datosSiguiente =
            await respuestaSiguiente.json();


        items =
            items.concat(
                datosSiguiente.value || []
            );


        siguiente =
            datosSiguiente["@odata.nextLink"];

    }


    return items;
}


/* =========================================================
   NORMALIZAR RESERVAS
========================================================= */

function normalizarReservas(items) {

    return items.map(
        item => {

            const fields =
                item.fields || {};


            return {

                id: item.id,

                Sala:
                    fields.Sala || "",

                FechaReserva:
                    fields.FechaReserva || "",

                HoraInicio:
                    fields.HoraInicio || "",

                HoraFin:
                    fields.HoraFin || "",

                Motivo:
                    fields.Motivo || "",

                Solicitante:
                    fields.Solicitante || "",

                CorreoSolicitante:
                    fields.CorreoSolicitante || "",

                Estado:
                    fields.Estado || "",

                FechaCreacion:
                    fields.FechaCreacion || "",

                CanceladoPor:
                    fields.CanceladoPor || "",

                FechaCancelacion:
                    fields.FechaCancelacion || "",

                IDReserva:
                    fields.IDReserva || "",

                BloqueInicio:
                    fields.BloqueInicio ?? "",

                BloqueFin:
                    fields.BloqueFin ?? ""

            };

        }
    );

}


/* =========================================================
   COMPROBAR ESTADO
========================================================= */

function esActiva(reserva) {

    return String(
        reserva.Estado || ""
    ).trim().toLowerCase() === "activa";

}


function esCancelada(reserva) {

    return String(
        reserva.Estado || ""
    ).trim().toLowerCase() === "cancelada";

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
                esActiva(reserva)
        );

}


/* =========================================================
   RESERVAS CANCELADAS DE HOY
========================================================= */

function obtenerReservasCanceladasHoy() {

    return obtenerReservasHoy()
        .filter(
            reserva =>
                esCancelada(reserva)
        );

}


/* =========================================================
   ACTUALIZAR ESTADÍSTICAS
========================================================= */

function actualizarEstadisticas() {

    const reservasHoy =
        obtenerReservasHoy();


    const activasHoy =
        reservasHoy.filter(
            reserva =>
                esActiva(reserva)
        );


    const canceladasHoy =
        reservasHoy.filter(
            reserva =>
                esCancelada(reserva)
        );


    /*
    ======================================
    TOTAL DE RESERVAS HOY
    ======================================
    */

    if (elementos.reservasHoy) {

        elementos.reservasHoy.textContent =
            reservasHoy.length;

    }


    /*
    ======================================
    RESERVAS ACTIVAS
    ======================================
    */

    if (elementos.reservasActivas) {

        elementos.reservasActivas.textContent =
            activasHoy.length;

    }


    /*
    ======================================
    RESERVAS CANCELADAS
    ======================================
    */

    if (elementos.reservasCanceladas) {

        elementos.reservasCanceladas.textContent =
            canceladasHoy.length;

    }


    /*
    ======================================
    TOTAL DE SALAS
    ======================================
    */

    if (elementos.totalSalas) {

        elementos.totalSalas.textContent =
            "3";

    }


    /*
    ======================================
    SALA 2
    ======================================
    */

    if (elementos.sala2Reservas) {

        elementos.sala2Reservas.textContent =
            reservasHoy.filter(
                reserva =>
                    reserva.Sala === "Sala 2"
            ).length;

    }


    /*
    ======================================
    SALA 3
    ======================================
    */

    if (elementos.sala3Reservas) {

        elementos.sala3Reservas.textContent =
            reservasHoy.filter(
                reserva =>
                    reserva.Sala === "Sala 3"
            ).length;

    }


    /*
    ======================================
    COMEDOR
    ======================================
    */

    if (elementos.comedorReservas) {

        elementos.comedorReservas.textContent =
            reservasHoy.filter(
                reserva =>
                    reserva.Sala === "Comedor"
            ).length;

    }


    /*
    ======================================
    TOTAL DE LA TABLA
    ======================================
    */

    if (elementos.totalReservasLabel) {

        elementos.totalReservasLabel.textContent =
            `${reservasHoy.length} reserva${reservasHoy.length === 1 ? "" : "s"}`;

    }

}


/* =========================================================
   OBTENER INICIALES
========================================================= */

function obtenerIniciales(nombre) {

    if (!nombre) {
        return "—";
    }


    const palabras =
        nombre
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
        palabras[palabras.length - 1][0]
    ).toUpperCase();

}


/* =========================================================
   GENERAR ESTADO VISUAL
========================================================= */

function generarEstado(estado) {

    const estadoNormalizado =
        String(
            estado || ""
        )
        .trim()
        .toLowerCase();


    if (estadoNormalizado === "cancelada") {

        return `
            <span class="estado estado-cancelada">
                <span class="estado-dot"></span>
                Cancelada
            </span>
        `;

    }


    return `
        <span class="estado estado-activa">
            <span class="estado-dot"></span>
            Activa
        </span>
    `;

}


/* =========================================================
   RENDERIZAR TABLA
========================================================= */

function renderizarReservas() {

    if (!elementos.tablaReservas) {
        return;
    }


    const tbody =
        elementos.tablaReservas;


    tbody.innerHTML = "";


    const reservasHoy =
        obtenerReservasHoy();


    /*
    ======================================
    ORDENAR POR HORA
    ======================================
    */

    const ordenadas =
        [...reservasHoy].sort(
            (a, b) =>
                String(a.HoraInicio || "")
                    .localeCompare(
                        String(b.HoraInicio || "")
                    )
        );


    /*
    ======================================
    EMPTY STATE
    ======================================
    */

    if (ordenadas.length === 0) {

        if (elementos.emptyState) {

            elementos.emptyState.style.display =
                "block";

        }

        return;

    }


    if (elementos.emptyState) {

        elementos.emptyState.style.display =
            "none";

    }


    /*
    ======================================
    CREAR FILAS
    ======================================
    */

    ordenadas.forEach(
        reserva => {

            const fila =
                document.createElement("tr");


            fila.innerHTML = `

                <td>
                    <div class="sala-cell">
                        ${reserva.Sala || "-"}
                    </div>
                </td>

                <td>
                    ${formatearFecha(reserva.FechaReserva)}
                </td>

                <td>
                    <strong>
                        ${reserva.HoraInicio || "-"}
                    </strong>
                    -
                    <strong>
                        ${reserva.HoraFin || "-"}
                    </strong>
                </td>

                <td>
                    <div class="solicitante-cell">

                        <div class="solicitante-avatar">
                            ${obtenerIniciales(reserva.Solicitante)}
                        </div>

                        <div>
                            <div class="solicitante-nombre">
                                ${reserva.Solicitante || "Sin registrar"}
                            </div>

                            ${
                                reserva.Motivo
                                    ? `
                                        <div class="solicitante-motivo">
                                            ${reserva.Motivo}
                                        </div>
                                      `
                                    : ""
                            }

                        </div>

                    </div>
                </td>

                <td>
                    ${generarEstado(reserva.Estado)}
                </td>

            `;


            tbody.appendChild(
                fila
            );

        }
    );

}


/* =========================================================
   OBTENER ESTADO DE UNA SALA EN TIEMPO REAL
========================================================= */

function obtenerEstadoSala(
    sala
) {

    const ahora =
        new Date();


    const horaActual =
        ahora.getHours() * 60 +
        ahora.getMinutes();


    const reservasSala =
        obtenerReservasActivasHoy()
            .filter(
                reserva =>
                    reserva.Sala === sala
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


                return (
                    inicio !== null &&
                    fin !== null &&
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
   CONVERTIR HORA A MINUTOS
========================================================= */

function convertirHoraAMinutos(hora) {

    if (!hora) {
        return null;
    }


    const partes =
        String(hora).split(":");


    if (partes.length < 2) {
        return null;
    }


    const horas =
        Number(partes[0]);


    const minutos =
        Number(partes[1]);


    if (
        Number.isNaN(horas) ||
        Number.isNaN(minutos)
    ) {

        return null;

    }


    return (
        horas * 60 +
        minutos
    );

}


/* =========================================================
   MOSTRAR ESTADO EN CONSOLA
========================================================= */

function mostrarEstadosSalas() {

    const salas = [
        "Sala 2",
        "Sala 3",
        "Comedor"
    ];


    console.log(
        "======================================"
    );

    console.log(
        "ESTADO ACTUAL DE LAS SALAS"
    );

    console.log(
        "======================================"
    );


    salas.forEach(
        sala => {

            const resultado =
                obtenerEstadoSala(sala);


            console.log(
                `${sala}: ${resultado.estado}`
            );


            if (resultado.reserva) {

                console.log(
                    `${resultado.reserva.HoraInicio} - ${resultado.reserva.HoraFin}`
                );

                console.log(
                    `Solicitante: ${resultado.reserva.Solicitante}`
                );

            }

        }
    );

}


/* =========================================================
   CARGAR RESERVAS
========================================================= */

async function cargarReservas() {

    try {

        /*
        ======================================
        MOSTRAR CARGANDO
        ======================================
        */

        console.log(
            "======================================"
        );

        console.log(
            "CARGANDO RESERVAS DE SALAS..."
        );

        console.log(
            "======================================"
        );


        /*
        ======================================
        TOKEN
        ======================================
        */

        const token =
            await obtenerToken();


        console.log(
            "TOKEN OBTENIDO"
        );


        /*
        ======================================
        SITE
        ======================================
        */

        const site =
            await obtenerSite(token);


        console.log(
            "SITE ENCONTRADO:",
            site.id
        );


        /*
        ======================================
        LISTA
        ======================================
        */

        const lista =
            await obtenerLista(
                token,
                site.id
            );


        console.log(
            "LISTA ENCONTRADA:",
            lista.displayName || lista.name
        );


        /*
        ======================================
        REGISTROS
        ======================================
        */

        const items =
            await obtenerRegistros(
                token,
                site.id,
                lista.id
            );


        console.log(
            `REGISTROS RECIBIDOS: ${items.length}`
        );


        /*
        ======================================
        NORMALIZAR
        ======================================
        */

        reservas =
            normalizarReservas(items);


        console.log(
            "RESERVAS NORMALIZADAS:",
            reservas
        );


        /*
        ======================================
        ACTUALIZAR INTERFAZ
        ======================================
        */

        actualizarEstadisticas();

        renderizarReservas();

        mostrarEstadosSalas();


        console.log(
            "======================================"
        );

        console.log(
            "RESERVAS CARGADAS CORRECTAMENTE"
        );

        console.log(
            "======================================"
        );

    }
    catch (error) {

        console.error(
            "======================================"
        );

        console.error(
            "ERROR CARGANDO RESERVAS"
        );

        console.error(
            "======================================"
        );

        console.error(error);


        /*
        ======================================
        MOSTRAR TABLA VACÍA
        ======================================
        */

        reservas = [];


        actualizarEstadisticas();

        renderizarReservas();

    }

}


/* =========================================================
   BOTÓN ACTUALIZAR
========================================================= */

function configurarRefresh() {

    if (!elementos.refreshButton) {
        return;
    }


    elementos.refreshButton.addEventListener(
        "click",
        async () => {

            const textoOriginal =
                elementos.refreshButton.innerHTML;


            elementos.refreshButton.disabled =
                true;


            try {

                elementos.refreshButton.innerHTML =
                    "Actualizando...";


                await cargarReservas();

            }
            finally {

                elementos.refreshButton.disabled =
                    false;


                elementos.refreshButton.innerHTML =
                    textoOriginal;

            }

        }
    );

}


/* =========================================================
   AUTO ACTUALIZACIÓN
========================================================= */

function iniciarActualizacionAutomatica() {

    /*
    Actualizar cada 60 segundos.
    */

    setInterval(
        async () => {

            await cargarReservas();

        },
        60000
    );

}


/* =========================================================
   INICIALIZAR
========================================================= */

async function iniciarSalas() {

    console.log(
        "======================================"
    );

    console.log(
        "ALFERZA LIVE OFFICE"
    );

    console.log(
        "MÓDULO DE SALAS"
    );

    console.log(
        "======================================"
    );


    actualizarFecha();

    configurarRefresh();

    await cargarReservas();

    iniciarActualizacionAutomatica();

}


/* =========================================================
   DOM READY
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        iniciarSalas();

    }
);