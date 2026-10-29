
/* =========================================
   PROTECCIÓN DE ACCESO
   ALFERZA LIVE OFFICE
========================================= */

(async function protegerPagina() {
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

    // Páginas públicas necesarias para autenticación.
    const paginasPublicas = [
        "login.html",
        "blank.html"
    ];

    if (paginasPublicas.includes(pagina)) {
        return;
    }

    // 1. Verificar la sesión de la aplicación.
    const sesion = sessionStorage.getItem("alferza_login");

    if (sesion !== "true") {
        sessionStorage.setItem("alferza_return_url", rutaActual);
        window.location.replace(LOGIN_URL);
        return;
    }

    // 2. Comprobar que se haya cargado el control central.
    if (!window.alferzaAccessReady) {
        console.error(
            "ALFERZA: no se ha cargado access-control.js."
        );

        mostrarErrorAcceso();
        return;
    }

    // 3. Esperar a que termine la validación de permisos.
    try {
        const resultado = await window.alferzaAccessReady;

        if (!resultado || resultado.autorizado !== true) {
            console.warn(
                "ALFERZA: acceso denegado o autorización no verificada."
            );
            return;
        }

        console.info(
            "ALFERZA: acceso autorizado correctamente."
        );

    } catch (error) {
        console.error(
            "ALFERZA: falló la comprobación de acceso.",
            error
        );

        mostrarErrorAcceso();
    }

    // Mostrar un mensaje si el control central no pudo cargarse.
    function mostrarErrorAcceso() {
        if (!document.body) {
            document.addEventListener(
                "DOMContentLoaded",
                mostrarErrorAcceso,
                { once: true }
            );
            return;
        }

        document.body.innerHTML = `
            <main style="
                min-height:100vh;
                display:flex;
                align-items:center;
                justify-content:center;
                padding:24px;
                box-sizing:border-box;
                background:#f4f6fa;
                font-family:Arial,sans-serif;
                color:#172033;
                text-align:center;
            ">
                <section style="
                    width:100%;
                    max-width:460px;
                    padding:36px 28px;
                    background:#fff;
                    border:1px solid #e1e6ef;
                    border-radius:16px;
                    box-sizing:border-box;
                ">
                    <h1 style="color:#00205c">
                        🔒 Acceso no verificado
                    </h1>

                    <p style="line-height:1.7;color:#596579">
                        No se pudo comprobar tu autorización.
                        Intenta iniciar sesión nuevamente.
                        Si el problema continúa, contacta con TI.
                    </p>

                    <button id="alferza-auth-login" type="button"
                        style="
                            margin-top:16px;
                            padding:12px 20px;
                            border:0;
                            border-radius:8px;
                            background:#00205c;
                            color:white;
                            cursor:pointer;
                        ">
                        Volver al inicio de sesión
                    </button>
                </section>
            </main>
        `;

        document
            .getElementById("alferza-auth-login")
            .addEventListener("click", function () {
                sessionStorage.removeItem("alferza_login");
                window.location.replace(LOGIN_URL);
            });
    }

})();
