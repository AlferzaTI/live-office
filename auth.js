/* =========================================
   PROTECCIÓN DE ACCESO
========================================= */

(function protegerPagina() {

    const sesion = sessionStorage.getItem("alferza_login");

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

    /*
     * Si no estamos en login.html,
     * guardamos la página actual.
     */
    if (!paginaActual.endsWith("/login.html")) {

        sessionStorage.setItem(
            "alferza_return_url",
            paginaActual
        );

    }

    /*
     * Mandamos al login.
     */
    window.location.replace("/login.html");

})();