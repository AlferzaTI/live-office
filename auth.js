/* =========================================
   PROTECCIÓN DE ACCESO
========================================= */

(function protegerPagina() {

    const sesion =
        sessionStorage.getItem("alferza_login");


    /*
     * Si ya tiene una sesión válida,
     * puede continuar normalmente.
     */

    if (sesion === "true") {

        return;

    }


    /*
     * Guardamos la página exacta que el usuario
     * estaba intentando abrir.
     */

    const paginaActual =
        window.location.pathname +
        window.location.search +
        window.location.hash;


    /*
     * No guardamos login.html como destino.
     */

    if (
        !paginaActual.endsWith("/login.html") &&
        !paginaActual.endsWith("/login")
    ) {

        sessionStorage.setItem(
            "alferza_return_url",
            paginaActual
        );

    }


    /*
     * Redirigimos al login.
     */

    window.location.replace(
        "/live-office/login.html"
    );

})();