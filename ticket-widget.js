document.addEventListener("DOMContentLoaded", function () {

    // Evitar que se cree dos veces
    if (document.querySelector(".ticket-widget")) {
        return;
    }

    const widget = document.createElement("div");

    widget.className = "ticket-widget";
    widget.title = "Registrar ticket";

    widget.innerHTML = `
        <img 
            src="img/ticket.gif" 
            alt="Registrar ticket"
        >
    `;

    document.body.appendChild(widget);

    widget.addEventListener("click", abrirTicketWidget);
});


function abrirTicketWidget() {

    alert("🎫 Aquí se abrirá el formulario de ticket");

}

