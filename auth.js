/* =========================================================
   ALFERZA LIVE OFFICE
   AUTENTICACIÓN GENERAL
========================================================= */

(function () {

    "use strict";


    /* =====================================================
       CONFIGURACIÓN
    ===================================================== */

    const MSAL_CONFIG = {

        auth: {

            clientId:
                "5d98417c-74a7-4fab-8f2c-41ac127be696",

            authority:
                "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",

            redirectUri:
                new URL(
                    "blank.html",
                    window.location.href
                ).href

        },

        cache: {

            cacheLocation:
                "sessionStorage",

            storeAuthStateInCookie:
                false

        }

    };


    /* =====================================================
       VALIDAR MSAL
    ===================================================== */

    if (
        typeof msal === "undefined"
    ) {

        console.error(
            "MSAL no está disponible."
        );

        return;

    }


    /* =====================================================
       INSTANCIA MSAL
    ===================================================== */

    const msalInstance =
        new msal.PublicClientApplication(
            MSAL_CONFIG
        );


    /*
     * Exponemos la instancia para que otros módulos
     * puedan reutilizar la misma sesión.
     */

    window.alferzaMsal =
        msalInstance;


    /* =====================================================
       INICIALIZAR AUTENTICACIÓN
    ===================================================== */

    async function inicializarAuth() {

        try {

            /*
             * Procesar respuesta del login si existe.
             */

            await msalInstance.handleRedirectPromise();


            /*
             * Obtener cuentas existentes.
             */

            const cuentas =
                msalInstance.getAllAccounts();


            console.log(
                "Cuentas MSAL encontradas:",
                cuentas
            );


            /*
             * No hay sesión de Microsoft.
             */

            if (
                !cuentas.length
            ) {

                protegerPagina();

                return null;

            }


            /*
             * Seleccionar la primera cuenta.
             */

            const cuenta =
                cuentas[0];


            msalInstance.setActiveAccount(
                cuenta
            );


            /*
             * Guardamos indicador de sesión
             * para mantener compatibilidad con
             * el resto de ALFERZA LIVE OFFICE.
             */

            sessionStorage.setItem(
                "alferza_login",
                "true"
            );


            console.log(
                "Cuenta seleccionada:",
                cuenta
            );


            return cuenta;

        }

        catch (error) {

            console.error(
                "Error inicializando autenticación:",
                error
            );


            sessionStorage.removeItem(
                "alferza_login"
            );


            protegerPagina();


            return null;

        }

    }


    /* =====================================================
       PROTEGER PÁGINA
       
       IMPORTANTE:
       Esto SOLO comprueba LOGIN.
       
       NO comprueba permisos de módulos.
    ===================================================== */

    function protegerPagina() {

        const paginaActual =
            window.location.pathname +
            window.location.search +
            window.location.hash;


        const esLogin =

            paginaActual ===
                "/live-office/login.html" ||

            paginaActual ===
                "/live-office/" ||

            paginaActual.endsWith(
                "/login.html"
            );


        if (
            esLogin
        ) {

            return;

        }


        sessionStorage.setItem(
            "alferza_return_url",
            paginaActual
        );


        window.location.replace(
            "/live-office/login.html"
        );

    }


    /* =====================================================
       OBTENER CUENTA ACTUAL
    ===================================================== */

    function obtenerCuentaActual() {

        const cuentas =
            msalInstance.getAllAccounts();


        if (
            !cuentas.length
        ) {

            return null;

        }


        const activa =
            msalInstance.getActiveAccount();


        if (
            activa
        ) {

            return activa;

        }


        msalInstance.setActiveAccount(
            cuentas[0]
        );


        return cuentas[0];

    }


    /* =====================================================
       OBTENER TOKEN GRAPH
    ===================================================== */

    async function obtenerTokenGraph(
        scopes = [
            "User.Read",
            "Sites.ReadWrite.All"
        ]
    ) {

        const cuenta =
            obtenerCuentaActual();


        if (!cuenta) {

            const error =
                new Error(
                    "No existe una sesión de Microsoft."
                );


            error.codigo =
                "SIN_SESION";


            throw error;

        }


        try {

            const resultado =
                await msalInstance.acquireTokenSilent({

                    scopes:
                        scopes,

                    account:
                        cuenta

                });


            return resultado.accessToken;

        }

        catch (error) {

            console.error(
                "Error obteniendo token Graph:",
                error
            );


            const nuevoError =
                new Error(
                    "No se pudo obtener el token de Microsoft."
                );


            nuevoError.codigo =
                "REQUIERE_INTERACCION";


            throw nuevoError;

        }

    }


    /* =====================================================
       FETCH GRAPH
    ===================================================== */

    async function graphFetch(
        url,
        opciones = {}
    ) {

        const token =
            await obtenerTokenGraph();


        const headers = {

            Authorization:
                `Bearer ${token}`,

            "Content-Type":
                "application/json",

            ...(opciones.headers || {})

        };


        const respuesta =
            await fetch(

                url,

                {

                    ...opciones,

                    headers

                }

            );


        if (!respuesta.ok) {

            const texto =
                await respuesta.text();


            throw new Error(

                `Graph ${respuesta.status}: ${texto}`

            );

        }


        if (
            respuesta.status ===
            204
        ) {

            return null;

        }


        const texto =
            await respuesta.text();


        if (!texto) {

            return null;

        }


        try {

            return JSON.parse(
                texto
            );

        }

        catch {

            return texto;

        }

    }


    /* =====================================================
       USUARIO ACTUAL
    ===================================================== */

    function obtenerCorreoCuentaActual() {

        const cuenta =
            obtenerCuentaActual();


        if (!cuenta) {

            return "";

        }


        return String(

            cuenta.username ||

            cuenta.idTokenClaims?.preferred_username ||

            cuenta.idTokenClaims?.email ||

            ""

        )
            .trim()
            .toLowerCase();

    }


    /* =====================================================
       EXPONER API GLOBAL
    ===================================================== */

    window.alferzaAuth = {

        msal:
            msalInstance,

        inicializar:
            inicializarAuth,

        obtenerCuenta:
            obtenerCuentaActual,

        obtenerToken:
            obtenerTokenGraph,

        graphFetch:
            graphFetch,

        obtenerCorreo:
            obtenerCorreoCuentaActual

    };


    /* =====================================================
       INICIO
    ===================================================== */

    document.addEventListener(

        "DOMContentLoaded",

        function () {

            inicializarAuth();

        }

    );

})();