
/* =========================================================
   ALFERZA LIVE OFFICE
   GESTIÓN DE SALAS
========================================================= */


/* =========================================================
   DATOS
========================================================= */

const reservas = [

    {
        sala: "Sala 2",
        fecha: "21/09/2026",
        inicio: "08:00",
        fin: "09:00",
        solicitante: "Patricia Condori",
        estado: "Activa"
    },

    {
        sala: "Sala 2",
        fecha: "21/09/2026",
        inicio: "17:00",
        fin: "17:30",
        solicitante: "Alvaro Durand",
        estado: "Activa"
    },

    {
        sala: "Sala 3",
        fecha: "21/09/2026",
        inicio: "09:30",
        fin: "10:30",
        solicitante: "Tyron",
        estado: "Activa"
    },

    {
        sala: "Comedor",
        fecha: "21/09/2026",
        inicio: "12:00",
        fin: "13:00",
        solicitante: "Pablo",
        estado: "Cancelada"
    }

];


/* =========================================================
   ELEMENTOS DOM
========================================================= */

const tbody =
    document.getElementById("tablaReservas");

const emptyState =
    document.getElementById("emptyState");

const reservasHoy =
    document.getElementById("reservasHoy");

const reservasActivas =
    document.getElementById("reservasActivas");

const reservasCanceladas =
    document.getElementById("reservasCanceladas");

const totalReservasLabel =
    document.getElementById("totalReservasLabel");

const sala2Reservas =
    document.getElementById("sala2Reservas");

const sala3Reservas =
    document.getElementById("sala3Reservas");

const comedorReservas =
    document.getElementById("comedorReservas");

const fechaActual =
    document.getElementById("fechaActual");


/* =========================================================
   UTILIDADES
========================================================= */


/**
 * Obtiene las iniciales de una persona.
 */
function obtenerIniciales(nombre) {

    const partes =
        nombre
            .trim()
            .split(/\s+/);

    if (partes.length === 1) {
        return partes[0]
            .substring(0, 2)
            .toUpperCase();
    }

    return (
        partes[0][0] +
        partes[1][0]
    ).toUpperCase();

}


/**
 * Genera el HTML correspondiente
 * al estado de una reserva.
 */
function generarEstado(estado) {

    if (estado === "Activa") {

        return `
            <span class="status-badge status-active">
                Activa
            </span>
        `;

    }

    return `
        <span class="status-badge status-cancelled">
            Cancelada
        </span>
    `;

}


/* =========================================================
   RENDERIZAR RESERVAS
========================================================= */

function renderizarReservas() {

    tbody.innerHTML = "";

    /*
     * Si no existen reservas,
     * mostramos el estado vacío.
     */

    if (reservas.length === 0) {

        emptyState.hidden = false;

        return;

    }

    emptyState.hidden = true;


    /*
     * Construimos todas las filas.
     */

    const filas = reservas.map(reserva => {

        const iniciales =
            obtenerIniciales(
                reserva.solicitante
            );

        return `
            <tr>

                <td>
                    ${reserva.sala}
                </td>

                <td>
                    ${reserva.fecha}
                </td>

                <td>
                    <span class="time-cell">
                        ${reserva.inicio}
                        —
                        ${reserva.fin}
                    </span>
                </td>

                <td>

                    <div class="applicant-cell">

                        <div class="avatar">
                            ${iniciales}
                        </div>

                        <span>
                            ${reserva.solicitante}
                        </span>

                    </div>

                </td>

                <td>
                    ${generarEstado(reserva.estado)}
                </td>

            </tr>
        `;

    }).join("");


    tbody.innerHTML = filas;

}


/* =========================================================
   CALCULAR ESTADÍSTICAS
========================================================= */

function actualizarEstadisticas() {

    let sala2 = 0;
    let sala3 = 0;
    let comedor = 0;

    let activas = 0;
    let canceladas = 0;


    reservas.forEach(reserva => {

        /*
         * Contador por sala
         */

        switch (reserva.sala) {

            case "Sala 2":
                sala2++;
                break;

            case "Sala 3":
                sala3++;
                break;

            case "Comedor":
                comedor++;
                break;

        }


        /*
         * Contador por estado
         */

        if (reserva.estado === "Activa") {
            activas++;
        }

        if (reserva.estado === "Cancelada") {
            canceladas++;
        }

    });


    /*
     * Estadísticas generales
     */

    reservasHoy.textContent =
        reservas.length;

    reservasActivas.textContent =
        activas;

    reservasCanceladas.textContent =
        canceladas;

    totalReservasLabel.textContent =
        reservas.length;


    /*
     * Estadísticas por sala
     */

    sala2Reservas.textContent =
        sala2;

    sala3Reservas.textContent =
        sala3;

    comedorReservas.textContent =
        comedor;

}


/* =========================================================
   FECHA ACTUAL
========================================================= */

function actualizarFecha() {

    const ahora = new Date();

    const opciones = {
        day: "2-digit",
        month: "short",
        year: "numeric"
    };

    let fecha =
        ahora.toLocaleDateString(
            "es-PE",
            opciones
        );

    /*
     * Normalizamos la fecha para
     * mantener un formato corporativo.
     */

    fecha =
        fecha
            .replace(".", "")
            .toUpperCase();

    fechaActual.textContent =
        fecha;

}


/* =========================================================
   INICIALIZACIÓN
========================================================= */

function iniciarSalas() {

    renderizarReservas();

    actualizarEstadisticas();

    actualizarFecha();

}


/* =========================================================
   EJECUTAR
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    iniciarSalas
);
