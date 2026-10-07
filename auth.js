/* =========================================
   PROTECCIÓN DE ACCESO
========================================= */

(function protegerPagina() {

    const sesion =
        sessionStorage.getItem("alferza_login");


    /*
     * Si ya existe una sesión válida,
     * dejamos continuar.
     */

    if (sesion === "true") {

        return;

    }


    /*
     * Guardamos la página que el usuario
     * intentaba abrir.
     */

    const paginaActual =
        window.location.pathname +
        window.location.search +
        window.location.hash;


    if (
        paginaActual !== "/live-office/login.html" &&
        paginaActual !== "/live-office/"
    ) {

        sessionStorage.setItem(
            "alferza_return_url",
            paginaActual
        );

    }


    /*
     * Mandamos al login.
     */

    window.location.replace(
        "/live-office/login.html"
    );

})();