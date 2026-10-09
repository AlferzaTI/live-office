
/* =========================================
   PROTECCIÓN DE ACCESO
   ALFERZA LIVE OFFICE
========================================= */

(function protegerPagina() {
    "use strict";

    const BASE_PATH = "/live-office/";
    const LOGIN_URL = BASE_PATH + "login.html";

    const rutaActual =
        window.location.pathname +
        window.location.search +
        window.location.hash;

    const pagina = window.location.pathname
        .split("/")
        .pop()
        .toLowerCase();

    // Páginas necesarias para iniciar o completar la autenticación.
    const paginasPublicas = [
        "login.html",
        "blank.html"
    ];

    if (paginasPublicas.includes(pagina)) {
        return;
    }

    // 1. Comprobar sesión de la aplicación.
    const sesion = sessionStorage.getItem("alferza_login");

    if (sesion !== "true") {
        sessionStorage.setItem("alferza_return_url", rutaActual);
        window.location.replace(LOGIN_URL);
        return;
    }

    // 2. Esperar a la autorización central si está instalada.
    // access-control.js debe cargarse antes que los scripts del módulo.
    if (window.alferzaAccessReady) {
        window.alferzaAccessReady.then(function (resultado) {
            if (!resultado || resultado.autorizado !== true) {
                console.warn(
                    "ALFERZA: acceso no autorizado o no verificado."
                );
            }
        }).catch(function (error) {
            console.error(
                "ALFERZA: falló la comprobación de acceso.",
                error
            );
        });
    }
})();