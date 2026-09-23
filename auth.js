/* =========================================
   PROTECCIÓN DE ACCESO
   ========================================= */

(function protegerPagina() {

    const sesion =
        sessionStorage.getItem("alferza_login");


    /*
     * Si NO existe una sesión válida,
     * enviamos al usuario al login.
     */

    if (sesion !== "true") {

        window.location.replace(
            "/live-office/login.html"
        );

        return;

    }

})();