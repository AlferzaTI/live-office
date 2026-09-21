import { obtenerToken } from "./salas-api.js";


/* =========================================================
   CONFIGURACIÓN
========================================================= */

const SITE_PATH =
    "https://graph.microsoft.com/v1.0/sites/alferzaholding-my.sharepoint.com:/personal/soporte1_alferza_pe";

const NOMBRE_LISTA = "ReservaSalas";

let reservas = [];


/* =========================================================
   ELEMENTOS DEL HTML
========================================================= */

const elementos = {
    totalSalas:
        document.getElementById("totalSalas"),

    reservasHoy:
        document.getElementById("reservasHoy"),

    reservasActivas:
        document.getElementById("reservasActivas"),

    reservasCanceladas:
        document.getElementById("reservasCanceladas"),

    sala2Reservas:
        document.getElementById("sala2Reservas"),

    sala3Reservas:
        document.getElementById("sala3Reservas"),

    comedorReservas:
        document.getElementById("comedorReservas"),

    fechaActual:
        document.getElementById("fechaActual"),

    tablaReservas:
        document.getElementById("tablaReservas"),

    emptyState:
        document.getElementById("emptyState"),

    totalReservasLabel:
        document.getElementById("totalReservasLabel"),

    refreshButton:
        document.getElementById("refreshSalas")
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
   FECHA DE RESERVA
========================================================= */

function obtenerFechaReservaISO(reserva) {

    if (!reserva) {
        return "";
    }

    if (!reserva.FechaReserva) {
        return "";
    }

    return String(
        reserva.FechaReserva
    ).substring(0, 10);
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
        return fechaISO;
    }

    return `${partes[2]}/${partes[1]}/${partes[0]}`;
}


/* =========================================================
   CONVERTIR HORA A MINUTOS
========================================================= */

function convertirHoraAMinutos(hora) {

    if (!hora) {
        return null;
    }

    const texto =
        String(hora).trim();

    const match =
        texto.match(
            /^(\d{1,2}):(\d{2})/
        );

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

    return (
        horas * 60 +
        minutos
    );
}


/* =========================================================
   FORMATEAR HORA
========================================================= */

function formatearHora(hora) {

    if (!hora) {
        return "--:--";
    }

    const texto =
        String(hora).trim();

    const match =
        texto.match(
            /^(\d{1,2}):(\d{2})/
        );

    if (!match) {
        return texto;
    }

    return (
        String(match[1]).padStart(2, "0") +
        ":" +
        match[2]
    );
}


/* =========================================================
   NORMALIZAR RESERVAS
========================================================= */

function normalizarReservas(items) {

    return items.map(item => {

        /*
         * IMPORTANTE
         *
         * Tus datos están llegando directamente
         * en el objeto:
         *
         * {
         *   Sala: "Sala 2",
         *   FechaReserva: "...",
         *   Eleccion: "Activa"
         * }
         *
         * Pero dependiendo de la respuesta de Graph,
         * también podrían venir dentro de:
         *
         * item.fields
         *
         * Por eso aceptamos ambas estructuras.
         */

        const fields =
            item.fields || item;


        return {

            id:
                item.id ||
                fields.id ||
                "",

            Sala:
                fields.Sala ||
                "",

            FechaReserva:
                fields.FechaReserva ||
                "",

            HoraInicio:
                fields.HoraInicio ||
                "",

            HoraFin:
                fields.HoraFin ||
                "",

            Motivo:
                fields.Motivo ||
                "",

            Solicitante:
                fields.Solicitante ||
                "",

            CorreoSolicitante:
                fields.CorreoSolicitante ||
                "",

            /*
             * CAMPO REAL DE SHAREPOINT:
             *
             * Eleccion
             */

            Estado:
                fields.Eleccion ||
                "",

            FechaCreacion:
                fields.FechaCreacion ||
                "",

            CanceladoPor:
                fields.CanceladoPor ||
                "",

            FechaCancelacion:
                fields.FechaCancelacion ||
                "",

            IDReserva:
                fields.IDReserva ||
                "",

            BloqueInicio:
                fields.BloqueInicio ??
                "",

            BloqueFin:
                fields.BloqueFin ??
                ""
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
   OBTENER TODOS LOS ITEMS
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
                `Error obteniendo items: ${respuesta.status} ${texto}`
            );
        }


        const datos =
            await respuesta.json();


        if (
            Array.isArray(
                datos.value
            )
        ) {

            todosLosItems.push(
                ...datos.value
            );
        }


        url =
            datos["@odata.nextLink"] ||
            null;
    }


    return todosLosItems;
}


/* =========================================================
   OBTENER RESERVAS
========================================================= */

async function obtenerReservas() {

    const token =
        await obtenerToken();


    console.log(
        "Token obtenido correctamente."
    );


    const sitio =
        await obtenerSitio(token);


    console.log(
        "Site ID:",
        sitio.id
    );


    const lista =
        await obtenerLista(
            token,
            sitio.id
        );


    console.log(
        "Lista:",
        lista.displayName
    );


    console.log(
        "Lista ID:",
        lista.id
    );


    const items =
        await obtenerItems(
            token,
            sitio.id,
            lista.id
        );


    console.log(
        "Cantidad de registros:",
        items.length
    );


    /*
     * Normalizamos los 180 registros.
     */

    const datos =
        normalizarReservas(items);


    /*
     * Mostrar algunas reservas para comprobar
     * que ahora sí se están leyendo correctamente.
     */

    console.log(
        "======================================"
    );

    console.log(
        "PRIMERAS RESERVAS NORMALIZADAS"
    );

    console.log(
        datos.slice(0, 5)
    );

    console.log(
        "======================================"
    );


    /*
     * Mostrar específicamente las reservas
     * de HOY.
     */

    const hoy =
        obtenerFechaActualISO();


    const reservasHoy =
        datos.filter(
            reserva =>
                obtenerFechaReservaISO(
                    reserva
                ) === hoy
        );


    console.log(
        "======================================"
    );

    console.log(
        "RESERVAS ENCONTRADAS PARA HOY:",
        hoy
    );

    console.log(
        reservasHoy
    );

    console.log(
        "TOTAL DE HOY:",
        reservasHoy.length
    );

    console.log(
        "======================================"
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
            obtenerFechaReservaISO(
                reserva
            ) === hoy
    );
}


/* =========================================================
   RESERVAS ACTIVAS DE HOY
========================================================= */

function obtenerReservasActivasHoy() {

    return obtenerReservasHoy()
        .filter(
            reserva =>
                String(
                    reserva.Estado
                )
                    .trim()
                    .toLowerCase() ===
                "activa"
        );
}


/* =========================================================
   ACTUALIZAR FECHA
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
                String(
                    reserva.Estado
                )
                    .trim()
                    .toLowerCase() ===
                "activa"
        );


    const canceladas =
        reservasHoy.filter(
            reserva =>
                String(
                    reserva.Estado
                )
                    .trim()
                    .toLowerCase() ===
                "cancelada"
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
        "======================================"
    );

    console.log(
        "CONTADORES"
    );

    console.log(
        "Reservas hoy:",
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

    console.log(
        "======================================"
    );
}


/* =========================================================
   CONTADORES POR SALA
========================================================= */

function actualizarContadoresSalas() {

    const reservasHoy =
        obtenerReservasHoy();


    function contarSala(nombreSala) {

        return reservasHoy.filter(
            reserva =>
                String(
                    reserva.Sala
                )
                    .trim()
                    .toLowerCase() ===
                nombreSala
                    .toLowerCase()
        ).length;
    }


    if (elementos.sala2Reservas) {

        elementos.sala2Reservas.textContent =
            contarSala(
                "Sala 2"
            );
    }


    if (elementos.sala3Reservas) {

        elementos.sala3Reservas.textContent =
            contarSala(
                "Sala 3"
            );
    }


    if (elementos.comedorReservas) {

        elementos.comedorReservas.textContent =
            contarSala(
                "Comedor"
            );
    }
}


/* =========================================================
   ESTADO ACTUAL DE UNA SALA
========================================================= */

function obtenerEstadoSala(
    nombreSala
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
                    String(
                        reserva.Sala
                    )
                        .trim()
                        .toLowerCase() ===
                    nombreSala
                        .toLowerCase()
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
   ACTUALIZAR ESTADO VISUAL DE LAS SALAS
========================================================= */

function actualizarEstadosVisualesSalas() {

    const tarjetas =
        document.querySelectorAll(
            ".sala-card"
        );


    tarjetas.forEach(
        tarjeta => {

            const titulo =
                tarjeta.querySelector(
                    "h3"
                );


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
                resultado.estado ===
                "Ocupada"
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
            }


            else {

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


/* =========================================================
   INICIALES DEL SOLICITANTE
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


    if (
        palabras.length === 1
    ) {

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
   RENDERIZAR TABLA
========================================================= */

function renderizarTabla() {

    if (!elementos.tablaReservas) {

        console.error(
            "No existe #tablaReservas en el HTML."
        );

        return;
    }


    const reservasHoy =
        obtenerReservasHoy();


    console.log(
        "Renderizando tabla.",
        "Reservas:",
        reservasHoy.length
    );


    elementos.tablaReservas.innerHTML =
        "";


    if (elementos.totalReservasLabel) {

        elementos.totalReservasLabel.textContent =
            reservasHoy.length;
    }


    /*
     * NO HAY RESERVAS
     */

    if (
        reservasHoy.length === 0
    ) {

        if (elementos.emptyState) {

            elementos.emptyState.hidden =
                false;
        }


        console.warn(
            "No hay reservas para:",
            obtenerFechaActualISO()
        );


        return;
    }


    /*
     * SÍ HAY RESERVAS
     */

    if (elementos.emptyState) {

        elementos.emptyState.hidden =
            true;
    }


    /*
     * Ordenar por hora.
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


                return (
                    horaA -
                    horaB
                );
            }
        );


    /*
     * Crear cada fila.
     */

    ordenadas.forEach(
        reserva => {

            const fila =
                document.createElement(
                    "tr"
                );


            const estado =
                String(
                    reserva.Estado
                )
                    .trim()
                    .toLowerCase();


            const esActiva =
                estado === "activa";


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

                        ${escaparHTML(
                            reserva.Sala ||
                            "Sin sala"
                        )}

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

                    <span
                        class="estado ${claseEstado}"
                    >

                        <i class="estado-dot"></i>

                        ${textoEstado}

                    </span>

                </td>

            `;


            elementos.tablaReservas
                .appendChild(
                    fila
                );
        }
    );


    console.log(
        "Tabla renderizada correctamente:",
        ordenadas.length,
        "filas."
    );
}


/* =========================================================
   CARGAR TODO
========================================================= */

async function cargarReservas() {

    try {

        console.log(
            "======================================"
        );

        console.log(
            "INICIANDO CARGA DE RESERVAS"
        );

        console.log(
            "======================================"
        );


        reservas =
            await obtenerReservas();


        console.log(
            "Reservas normalizadas:",
            reservas.length
        );


        actualizarFecha();

        actualizarContadores();

        actualizarContadoresSalas();

        renderizarTabla();

        actualizarEstadosVisualesSalas();


        console.log(
            "======================================"
        );

        console.log(
            "CARGA FINALIZADA CORRECTAMENTE"
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
            error
        );

        console.error(
            "======================================"
        );


        if (
            elementos.tablaReservas
        ) {

            elementos.tablaReservas.innerHTML = `

                <tr>

                    <td
                        colspan="5"
                        style="
                            text-align:center;
                            padding:30px;
                        "
                    >

                        Error al cargar las reservas.

                    </td>

                </tr>

            `;
        }
    }
}


/* =========================================================
   INICIO
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        await cargarReservas();


        /*
         * Actualizar automáticamente
         * cada 60 segundos.
         */

        setInterval(
            async () => {

                await cargarReservas();

            },
            60000
        );


        /*
         * Si existe un botón
         * #refreshSalas
         */

        if (
            elementos.refreshButton
        ) {

            elementos.refreshButton
                .addEventListener(
                    "click",
                    cargarReservas
                );
        }
    }
);