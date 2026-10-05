
/* =========================================================
   WIDGET GLOBAL DE TICKETS
   ========================================================= */


/* Evitar que se cree dos veces */

if (!document.querySelector(".ticket-widget")) {


    /* ==========================================
       CREAR WIDGET
       ========================================== */

    const widget = document.createElement("div");

    widget.className = "ticket-widget";

    widget.title = "Registrar ticket";


    /* ==========================================
       CREAR IMAGEN
       ========================================== */

    const img = document.createElement("img");

    img.alt = "Registrar ticket";


    /*
     * Obtiene automáticamente la ubicación
     * de este archivo:
     *
     * /js/ticket-widget.js
     *
     * y desde ahí encuentra:
     *
     * /img/ticket.gif
     */

    const scriptActual =
        document.currentScript;

    const rutaScript =
        new URL(
            scriptActual.src,
            window.location.href
        );


    const rutaImagen =
        new URL(
            "../img/ticket.gif",
            rutaScript
        );


    img.src =
        rutaImagen.href;


    /* ==========================================
       INSERTAR
       ========================================== */

    widget.appendChild(img);

    document.body.appendChild(widget);


    /* ==========================================
       CLICK
       ========================================== */

    widget.addEventListener(
        "click",
        abrirTicketWidget
    );

}


/* =========================================================
   ABRIR FORMULARIO
   ========================================================= */

function abrirTicketWidget() {

    /*
     * Por ahora solamente comprobamos
     * que el botón funciona.
     *
     * Aquí posteriormente colocaremos
     * el formulario real.
     */

    alert(
        "🎫 Abrir formulario de ticket"
    );

}

