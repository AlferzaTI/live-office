const reservas = [
    {
        sala: "Sala 2",
        usuario: "Patricia Condori",
        inicio: "08:00",
        fin: "09:00"
    },
    {
        sala: "Sala 3",
        usuario: "Luis Barrionuevo",
        inicio: "10:00",
        fin: "11:00"
    },
    {
        sala: "Comedor",
        usuario: "Alvaro Durand",
        inicio: "12:00",
        fin: "13:00"
    }
];

const tbody = document.getElementById("tablaReservas");

reservas.forEach(r => {

    tbody.innerHTML += `
        <tr>
            <td>${r.sala}</td>
            <td>${r.usuario}</td>
            <td>${r.inicio}</td>
            <td>${r.fin}</td>
        </tr>
    `;

});
