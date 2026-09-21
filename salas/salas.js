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

const tbody = document.getElementById("tablaReservas");

let sala2 = 0;
let sala3 = 0;
let comedor = 0;

let activas = 0;
let canceladas = 0;

reservas.forEach(r => {

    tbody.innerHTML += `
    <tr>
        <td>${r.sala}</td>
        <td>${r.fecha}</td>
        <td>${r.inicio}</td>
        <td>${r.fin}</td>
        <td>${r.solicitante}</td>
        <td>${r.estado}</td>
    </tr>
    `;

    if (r.sala === "Sala 2") sala2++;
    if (r.sala === "Sala 3") sala3++;
    if (r.sala === "Comedor") comedor++;

    if (r.estado === "Activa") activas++;
    if (r.estado === "Cancelada") canceladas++;

});

document.getElementById("reservasHoy").textContent = reservas.length;
document.getElementById("reservasActivas").textContent = activas;
document.getElementById("reservasCanceladas").textContent = canceladas;

document.getElementById("sala2Reservas").textContent =
`${sala2} Reservas`;

document.getElementById("sala3Reservas").textContent =
`${sala3} Reservas`;

document.getElementById("comedorReservas").textContent =
`${comedor} Reservas`;