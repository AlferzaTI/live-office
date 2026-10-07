/* =========================================================
   ALFERZA LIVE OFFICE
   SIDEBAR CENTRALIZADO
========================================================= */

(function () {

    /* =====================================================
       MENÚ
    ====================================================== */

    const MENU_ITEMS = [

        {
            nombre: "Oficina",
            icono: "🏢",
            ruta: "/live-office/index.html"
        },

        {
            nombre: "Personal",
            icono: "👥",
            ruta: "/live-office/personal.html"
        },

        {
            nombre: "Reservas",
            icono: "📝",
            ruta: "/live-office/reserva.html"
        },

        {
            nombre: "Salas",
            icono: "📅",
            ruta: "/live-office/salas/salas.html"
        },

        {
            nombre: "Comunicados",
            icono: "📢",
            ruta: "/live-office/comunicados.html"
        },

        {
            nombre: "Seguridad",
            icono: "🔐",
            ruta: "/live-office/seguridad.html"
        },

        {
            nombre: "Infraestructura",
            icono: "🏗️",
            ruta: "/live-office/infraestructura.html"
        },

        {
            nombre: "Tickets",
            icono: "🎫",
            ruta: "/live-office/tickets.html"
        },

        {
            nombre: "Permisos",
            icono: "🔐",
            ruta: "/live-office/permisos.html"
        },

        {
            nombre: "Configuración",
            icono: "⚙️",
            ruta: "/live-office/configuracion.html"
        }

    ];


    /* =====================================================
       OBTENER RUTA ACTUAL
    ====================================================== */

    function obtenerRutaActual() {

        return window.location.pathname
            .replace(/\/+$/, "");

    }


    /* =====================================================
       NORMALIZAR RUTA
    ====================================================== */

    function normalizarRuta(
        ruta
    ) {

        return String(ruta || "")
            .replace(/\/+$/, "");

    }


    /* =====================================================
       CREAR SIDEBAR
    ====================================================== */

    function crearSidebar(
        sidebar
    ) {

        if (!sidebar) {

            return;

        }


        const rutaActual =
            obtenerRutaActual();


        sidebar.innerHTML = `

            <h2>
                MENÚ
            </h2>

            <ul>

                ${MENU_ITEMS.map(

                    item => {

                        const rutaMenu =
                            normalizarRuta(
                                item.ruta
                            );


                        const activo =
                            rutaActual ===
                            rutaMenu;


                        return `

                            <li
                                class="${activo ? "active" : ""}"
                                data-ruta="${item.ruta}"
                            >
                                ${item.icono}
                                ${item.nombre}
                            </li>

                        `;

                    }

                ).join("")}

            </ul>

        `;


        /* =================================================
           EVENTOS
        ================================================== */

        sidebar
            .querySelectorAll(
                "li[data-ruta]"
            )
            .forEach(

                item => {

                    item.addEventListener(

                        "click",

                        function () {

                            const ruta =
                                this.dataset.ruta;


                            if (
                                ruta
                            ) {

                                window.location.href =
                                    ruta;

                            }

                        }

                    );

                }

            );

    }


    /* =====================================================
       INICIALIZAR
    ====================================================== */

    function inicializarSidebar() {

        const sidebars =
            document.querySelectorAll(
                ".sidebar"
            );


        sidebars.forEach(

            sidebar => {

                crearSidebar(
                    sidebar
                );

            }

        );

    }


    /* =====================================================
       DOM
    ====================================================== */

    if (
        document.readyState ===
        "loading"
    ) {

        document.addEventListener(
            "DOMContentLoaded",
            inicializarSidebar
        );

    }

    else {

        inicializarSidebar();

    }


})();