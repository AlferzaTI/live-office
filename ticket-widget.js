document.body.insertAdjacentHTML("beforeend", `
    <div class="ticket-widget" onclick="abrirTicketWidget()">
        <img src="img/ticket.gif" alt="Registrar ticket">
    </div>
`);

function abrirTicketWidget() {
    alert("Aquí se abrirá el formulario de ticket");
}