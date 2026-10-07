/* =========================================================
   ALFERZA LIVE OFFICE
   AUTENTICACIÓN GENERAL + PERMISOS POR MÓDULO

   Para proteger una página, agregar en el <body>:
        <body data-modulo="Oficina">

   Si la página no tiene data-modulo (login, index, etc.),
   solo se comprueba que exista sesión de Microsoft.
========================================================= */

(function () {

    "use strict";


    /* =====================================================
       CONFIGURACIÓN MSAL
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
       CONFIGURACIÓN DE PERMISOS
       (debe coincidir con permisos.js)
    ===================================================== */

    const SHAREPOINT_HOST =
        "alferzaholding-my.sharepoint.com";

    const SHAREPOINT_SITE_PATH =
        "/personal/soporte1_alferza_pe";

    const PERMISOS_LIST_NAME =
        "PermisosTI";

    const GRAPH =
        "https://graph.microsoft.com/v1.0";

    const GRUPO_TI =
        "ti";

    const USUARIOS_TI_INICIALES = [
        "soporte1@alferza.pe"
    ];

    const MODULOS = [
        "Oficina",
        "Personal",
        "Reservas",
        "Salas",
        "Comunicados",
        "Seguridad",
        "Infraestructura",
        "Tickets",
        "Configuracion"
    ];

    const PAGINA_DENEGADA =
        "/live-office/index.html";


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
       UTILIDADES
    ===================================================== */

    function normalizarTexto(texto) {

        return String(texto || "")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .trim()
            .toLowerCase();

    }


    function buscarColumna(columnas, nombres) {

        const objetivos =
            nombres.map(normalizarTexto);

        let columna =
            columnas.find(
                c => objetivos.includes(
                    normalizarTexto(c.displayName)
                )
            );

        if (!columna) {

            columna =
                columnas.find(
                    c => objetivos.includes(
                        normalizarTexto(c.name)
                    )
                );

        }

        return columna
            ? columna.name
            : null;

    }


    function esTIInicial(correo) {

        return USUARIOS_TI_INICIALES
            .map(normalizarTexto)
            .includes(normalizarTexto(correo));

    }


    /* =====================================================
       PROTEGER PÁGINA (SOLO LOGIN)
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
       LEER PERMISOSTI (CACHE EN MEMORIA)
    ===================================================== */

    let cacheRegistros = null;


    async function cargarRegistrosPermisos() {

        if (cacheRegistros) {

            return cacheRegistros;

        }


        const sitio =
            await graphFetch(
                `${GRAPH}/sites/${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`
            );


        const filtro =
            encodeURIComponent(
                `displayName eq '${PERMISOS_LIST_NAME}'`
            );


        const listas =
            await graphFetch(
                `${GRAPH}/sites/${sitio.id}/lists?$filter=${filtro}`
            );


        const lista =
            listas && listas.value
                ? listas.value[0]
                : null;


        if (!lista) {

            throw new Error(
                "No se encontró la lista PermisosTI."
            );

        }


        const respColumnas =
            await graphFetch(
                `${GRAPH}/sites/${sitio.id}/lists/${lista.id}/columns`
            );


        const columnas =
            respColumnas.value || [];


        const colCorreo =
            buscarColumna(columnas, ["UsuarioCorreo"]);

        const colGrupo =
            buscarColumna(columnas, ["Grupo"]);


        const colModulos = {};

        MODULOS.forEach(modulo => {

            colModulos[modulo] =
                buscarColumna(columnas, [modulo]);

        });


        if (!colCorreo) {

            throw new Error(
                "PermisosTI no tiene la columna UsuarioCorreo."
            );

        }


        /*
         * Leer todos los registros (con paginación).
         */

        let url =
            `${GRAPH}/sites/${sitio.id}/lists/${lista.id}/items?$expand=fields&$top=999`;

        const items = [];

        while (url) {

            const resp =
                await graphFetch(url);

            items.push(
                ...((resp && resp.value) || [])
            );

            url =
                (resp && resp["@odata.nextLink"]) || null;

        }


        const mapa = {};

        items.forEach(item => {

            const campos =
                item.fields || {};

            const correo =
                normalizarTexto(campos[colCorreo]);

            if (!correo) {

                return;

            }


            const modulos = new Set();

            MODULOS.forEach(modulo => {

                const col = colModulos[modulo];

                if (col && campos[col] === true) {

                    modulos.add(modulo);

                }

            });


            mapa[correo] = {

                grupo:
                    normalizarTexto(
                        colGrupo ? campos[colGrupo] : ""
                    ),

                modulos:
                    modulos

            };

        });


        cacheRegistros = mapa;

        return mapa;

    }


    /* =====================================================
       OBTENER PERMISOS DEL USUARIO
    ===================================================== */

    async function obtenerPermisosUsuario(correo) {

        /*
         * La cuenta TI inicial tiene acceso total
         * y no depende de leer la lista.
         */

        if (esTIInicial(correo)) {

            return {

                modulos:
                    new Set(MODULOS)

            };

        }


        const registros =
            await cargarRegistrosPermisos();


        const registro =
            registros[normalizarTexto(correo)];


        /*
         * Usuario no registrado: sin acceso.
         */

        if (!registro) {

            return {

                modulos:
                    new Set()

            };

        }


        /*
         * Grupo TI: acceso total.
         */

        if (registro.grupo === GRUPO_TI) {

            return {

                modulos:
                    new Set(MODULOS)

            };

        }


        return {

            modulos:
                registro.modulos

        };

    }


    /* =====================================================
       VERIFICAR ACCESO AL MÓDULO DE LA PÁGINA
       Si no hay permiso o no se puede leer PermisosTI,
       NO se concede acceso (falla cerrado).
    ===================================================== */

    async function verificarAccesoModulo(
        modulo,
        cuenta
    ) {

        const correo =
            String(

                cuenta.username ||

                cuenta.idTokenClaims?.preferred_username ||

                cuenta.idTokenClaims?.email ||

                ""

            ).trim().toLowerCase();


        try {

            const permisos =
                await obtenerPermisosUsuario(correo);


            if (
                permisos.modulos.has(modulo)
            ) {

                return true;

            }


            console.warn(
                `Sin acceso al módulo ${modulo}:`,
                correo
            );

            return false;

        }

        catch (error) {

            console.error(
                "Error verificando permisos:",
                error
            );

            return false;

        }

    }


    /* =====================================================
       MENÚ: OCULTAR ENLACES SIN PERMISO
       (enlaces con data-modulo-link="Nombre")
    ===================================================== */

    async function filtrarMenuPorPermisos(cuenta) {

        const correo =
            String(
                cuenta.username || ""
            ).trim().toLowerCase();


        try {

            const permisos =
                await obtenerPermisosUsuario(correo);


            document
                .querySelectorAll("[data-modulo-link]")
                .forEach(enlace => {

                    if (
                        !permisos.modulos.has(
                            enlace.dataset.moduloLink
                        )
                    ) {

                        enlace.style.display =
                            "none";

                    }

                });

        }

        catch (error) {

            console.error(
                "No se pudo filtrar el menú:",
                error
            );

        }

    }


    /* =====================================================
       INICIALIZAR AUTENTICACIÓN
    ===================================================== */

    async function inicializarAuth() {

        const modulo =
            document.body
                ? document.body.dataset.modulo
                : null;


        /*
         * Ocultar contenido mientras se verifican permisos.
         */

        if (modulo) {

            document.documentElement.style.visibility =
                "hidden";

        }


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
             * Indicador de sesión para el resto de la app.
             */

            sessionStorage.setItem(
                "alferza_login",
                "true"
            );


            /*
             * CONTROL DE PERMISOS POR MÓDULO
             */

            if (modulo) {

                const autorizado =
                    await verificarAccesoModulo(
                        modulo,
                        cuenta
                    );


                if (!autorizado) {

                    window.location.replace(
                        `${PAGINA_DENEGADA}?denegado=${encodeURIComponent(modulo)}`
                    );

                    return null;

                }


                await filtrarMenuPorPermisos(
                    cuenta
                );

                document.documentElement.style.visibility =
                    "";

            }


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