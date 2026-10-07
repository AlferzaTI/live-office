/* =========================================================
   PROTECCIÓN Y CONTROL DE PERMISOS
   ALFERZA LIVE OFFICE
========================================================= */

(function protegerPagina () {

    /* =========================================================
       CONFIGURACIÓN GENERAL
    ========================================================= */

    const SESION_KEY =
        "alferza_login";

    const RETURN_KEY =
        "alferza_return_url";

    const PERMISOS_KEY =
        "alferza_permisos";

    const LOGIN_PATH =
        "/live-office/login.html";

    const ROOT_PATH =
        "/live-office/";


    /* =========================================================
       SHAREPOINT
    ========================================================= */

    const SHAREPOINT_HOST =
        "alferzaholding-my.sharepoint.com";

    const SHAREPOINT_SITE_PATH =
        "/personal/soporte1_alferza_pe";

    const PERMISOS_LIST_NAME =
        "PermisosTI";


    /* =========================================================
       CUENTA TI INICIAL
       ESTA CUENTA SIEMPRE TIENE ACCESO TOTAL
    ========================================================= */

    const USUARIO_TI_INICIAL =
        "soporte1@alferza.pe";


    /* =========================================================
       MAPEO DE PÁGINAS
    ========================================================= */

    const PAGINAS_PERMISOS = {

        "index.html":
            "Oficina",

        "personal.html":
            "Personal",

        "reserva.html":
            "Reservas",

        "reservas.html":
            "Reservas",

        "salas/salas.html":
            "Salas",

        "salas.html":
            "Salas",

        "comunicados.html":
            "Comunicados",

        "seguridad.html":
            "Seguridad",

        "infraestructura.html":
            "Infraestructura",

        "tickets.html":
            "Tickets",

        "permisos.html":
            "Permisos",

        "configuracion.html":
            "Configuracion"

    };


    /* =========================================================
       CONFIGURACIÓN MSAL
    ========================================================= */

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


    /* =========================================================
       COMPROBAR SESIÓN BÁSICA
    ========================================================= */

    const sesion =
        sessionStorage.getItem(
            SESION_KEY
        );

    /*
     * Si no existe la marca de sesión,
     * enviamos al login.
     *
     * La validación real de Microsoft
     * se hace posteriormente mediante MSAL.
     */

    if (
        sesion !== "true"
    ) {

        guardarReturnUrl();

        window.location.replace(
            LOGIN_PATH
        );

        return;

    }


    /* =========================================================
       GUARDAR URL DE RETORNO
    ========================================================= */

    function guardarReturnUrl () {

        const paginaActual =
            window.location.pathname +
            window.location.search +
            window.location.hash;

        if (

            paginaActual !== LOGIN_PATH &&

            paginaActual !== ROOT_PATH

        ) {

            sessionStorage.setItem(

                RETURN_KEY,

                paginaActual

            );

        }

    }


    /* =========================================================
       INICIO
    ========================================================= */

    if (
        document.readyState ===
        "loading"
    ) {

        document.addEventListener(

            "DOMContentLoaded",

            iniciarControlPermisos

        );

    }

    else {

        iniciarControlPermisos();

    }


    /* =========================================================
       INICIAR CONTROL
    ========================================================= */

    async function iniciarControlPermisos () {

        console.log(
            "========================================="
        );

        console.log(
            "ALFERZA LIVE OFFICE - CONTROL DE PERMISOS"
        );

        console.log(
            "========================================="
        );

        console.log(
            "Página actual:",
            window.location.pathname
        );


        /* -----------------------------------------
           Aplicar temporalmente permisos guardados
        ----------------------------------------- */

        aplicarPermisosGuardados();


        try {

            const permisos =
                await obtenerPermisosUsuario();


            console.log(
                "PERMISOS OBTENIDOS:",
                permisos
            );


            /* -----------------------------------------
               Guardar permisos actuales
            ----------------------------------------- */

            sessionStorage.setItem(

                PERMISOS_KEY,

                JSON.stringify(
                    permisos
                )

            );


            /* -----------------------------------------
               Aplicar menú
            ----------------------------------------- */

            aplicarPermisosMenu(
                permisos
            );


            /* -----------------------------------------
               Validar página actual
            ----------------------------------------- */

            validarPaginaActual(
                permisos
            );

        }

        catch (error) {

            console.error(
                "========================================="
            );

            console.error(
                "ERROR OBTENIENDO PERMISOS"
            );

            console.error(
                error
            );

            console.error(
                "========================================="
            );


            /*
             * Si el error es de sesión,
             * no debemos continuar utilizando
             * permisos antiguos.
             */

            if (
                esErrorSesion(
                    error
                )
            ) {

                sessionStorage.removeItem(
                    SESION_KEY
                );

                sessionStorage.removeItem(
                    PERMISOS_KEY
                );

                guardarReturnUrl();

                window.location.replace(
                    LOGIN_PATH
                );

                return;

            }


            /*
             * Si Graph falla por otro motivo,
             * utilizamos los permisos guardados.
             */

            const permisosGuardados =
                obtenerPermisosGuardados();


            if (
                permisosGuardados
            ) {

                console.warn(
                    "Se utilizarán los permisos guardados en sesión."
                );


                aplicarPermisosMenu(
                    permisosGuardados
                );


                validarPaginaActual(
                    permisosGuardados
                );


                return;

            }


            /*
             * Si nunca existieron permisos válidos,
             * bloqueamos el acceso.
             */

            console.warn(
                "No existen permisos almacenados."
            );


            const permisosDenegados =
                crearPermisosDenegados();


            aplicarPermisosMenu(
                permisosDenegados
            );


            validarPaginaActual(
                permisosDenegados
            );

        }

    }


    /* =========================================================
       OBTENER PERMISOS GUARDADOS
    ========================================================= */

    function obtenerPermisosGuardados () {

        try {

            const almacenados =
                sessionStorage.getItem(
                    PERMISOS_KEY
                );


            if (
                !almacenados
            ) {

                return null;

            }


            const permisos =
                JSON.parse(
                    almacenados
                );


            if (

                permisos &&

                permisos.permisos

            ) {

                return permisos;

            }

        }

        catch (error) {

            console.warn(
                "No se pudieron leer permisos guardados:",
                error
            );

        }


        return null;

    }


    /* =========================================================
       APLICAR PERMISOS GUARDADOS
    ========================================================= */

    function aplicarPermisosGuardados () {

        const permisos =
            obtenerPermisosGuardados();


        if (
            !permisos
        ) {

            return;

        }


        console.log(
            "Permisos guardados encontrados:",
            permisos
        );


        aplicarPermisosMenu(
            permisos
        );

    }


    /* =========================================================
       OBTENER MSAL
    ========================================================= */

    function obtenerMSAL () {

        if (
            typeof msal !== "undefined"
        ) {

            return msal;

        }

        return null;

    }


    /* =========================================================
       ESPERAR MSAL
    ========================================================= */

    async function esperarMSAL () {

        const maxIntentos =
            50;


        for (

            let intento = 0;

            intento < maxIntentos;

            intento++

        ) {

            if (
                typeof msal !== "undefined"
            ) {

                return msal;

            }


            await new Promise(

                resolve =>

                    setTimeout(

                        resolve,

                        100

                    )

            );

        }


        throw new Error(
            "MSAL no está disponible."
        );

    }


    /* =========================================================
       CREAR MSAL
    ========================================================= */

    let instanciaMSAL =
        null;


    async function crearMSAL () {

        if (
            instanciaMSAL
        ) {

            return instanciaMSAL;

        }


        const libreria =
            await esperarMSAL();


        instanciaMSAL =
            new libreria.PublicClientApplication(
                MSAL_CONFIG
            );


        return instanciaMSAL;

    }


    /* =========================================================
       OBTENER CUENTA ACTUAL
    ========================================================= */

    async function obtenerCuentaActual () {

        const instancia =
            await crearMSAL();


        const cuentas =
            instancia.getAllAccounts();


        if (

            !cuentas ||

            !cuentas.length

        ) {

            const error =
                new Error(
                    "No existe una cuenta de Microsoft."
                );


            error.codigo =
                "SIN_SESION";


            throw error;

        }


        let cuenta =
            instancia.getActiveAccount();


        if (
            !cuenta
        ) {

            cuenta =
                cuentas[0];


            instancia.setActiveAccount(
                cuenta
            );

        }


        console.log(
            "Cuenta Microsoft:",
            cuenta
        );


        return {

            instancia:
                instancia,

            cuenta:
                cuenta

        };

    }


    /* =========================================================
       OBTENER TOKEN
       
       Permite solicitar diferentes scopes.
       Esto es importante para módulos como:
       
       - Permisos
       - Tickets
       - Seguridad
       - SharePoint
    ========================================================= */

    async function obtenerToken (

        instancia,

        cuenta,

        scopes = [

            "User.Read",

            "Sites.ReadWrite.All"

        ]

    ) {

        try {

            const resultado =
                await instancia.acquireTokenSilent({

                    scopes:
                        scopes,

                    account:
                        cuenta

                });


            return resultado.accessToken;

        }

        catch (error) {

            console.error(
                "No se pudo obtener token silenciosamente:",
                error
            );


            /*
             * Indicamos que se requiere
             * interacción con Microsoft.
             *
             * El módulo que solicite el token
             * podrá decidir si utiliza popup
             * o redirect.
             */

            const nuevoError =
                new Error(
                    "Se requiere autenticación adicional de Microsoft."
                );


            nuevoError.codigo =
                "REQUIERE_INTERACCION";


            nuevoError.originalError =
                error;


            throw nuevoError;

        }

    }


    /* =========================================================
       GRAPH FETCH
    ========================================================= */

    async function graphFetch (

        url,

        token

    ) {

        console.log(
            "Graph GET:",
            url
        );


        const respuesta =
            await fetch(

                url,

                {

                    method:
                        "GET",

                    headers: {

                        Authorization:
                            `Bearer ${token}`,

                        Accept:
                            "application/json"

                    }

                }

            );


        if (
            !respuesta.ok
        ) {

            const texto =
                await respuesta.text();


            const error =
                new Error(

                    `Graph ${respuesta.status}: ${texto}`

                );


            error.status =
                respuesta.status;


            throw error;

        }


        const texto =
            await respuesta.text();


        if (
            !texto
        ) {

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


    /* =========================================================
       OBTENER PERMISOS DEL USUARIO
    ========================================================= */

    async function obtenerPermisosUsuario () {

        const {

            instancia,

            cuenta

        } =
            await obtenerCuentaActual();


        const token =
            await obtenerToken(

                instancia,

                cuenta,

                [

                    "User.Read",

                    "Sites.ReadWrite.All"

                ]

            );


        /* -----------------------------------------
           OBTENER CORREO
        ----------------------------------------- */

        const correo =
            normalizarCorreo(

                cuenta.username ||

                cuenta.idTokenClaims?.preferred_username ||

                cuenta.idTokenClaims?.email ||

                ""

            );


        console.log(
            "Correo detectado:",
            correo
        );


        if (
            !correo
        ) {

            throw new Error(
                "No se pudo determinar el correo del usuario."
            );

        }


        /* -----------------------------------------
           VALIDAR DOMINIO ALFERZA
        ----------------------------------------- */

        if (
            !correo.endsWith(
                "@alferza.pe"
            )
        ) {

            const error =
                new Error(
                    "La cuenta no pertenece a ALFERZA."
                );


            error.codigo =
                "DOMINIO_NO_AUTORIZADO";


            throw error;

        }


        /* -----------------------------------------
           CUENTA TI INICIAL
        ----------------------------------------- */

        if (

            correo ===

            normalizarCorreo(
                USUARIO_TI_INICIAL
            )

        ) {

            console.log(
                "Usuario TI inicial detectado."
            );


            return crearPermisosTI(

                correo,

                cuenta.name ||
                correo

            );

        }


        /* -----------------------------------------
           SITIO SHAREPOINT
        ----------------------------------------- */

        const sitioUrl =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;


        const sitio =
            await graphFetch(

                sitioUrl,

                token

            );


        console.log(
            "Sitio SharePoint:",
            sitio
        );


        if (

            !sitio ||

            !sitio.id

        ) {

            throw new Error(
                "No se pudo obtener el sitio de SharePoint."
            );

        }


        /* -----------------------------------------
           OBTENER LISTAS
        ----------------------------------------- */

        const listasUrl =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${sitio.id}/lists` +

            `?$select=id,name,displayName&$top=200`;


        const listas =
            await graphFetch(

                listasUrl,

                token

            );


        console.log(
            "Listas SharePoint:",
            listas
        );


        const objetivoLista =
            normalizarTexto(
                PERMISOS_LIST_NAME
            );


        const lista =

            (

                listas.value || []

            ).find(

                item => {

                    const displayName =
                        normalizarTexto(
                            item.displayName
                        );


                    const name =
                        normalizarTexto(
                            item.name
                        );


                    return (

                        displayName ===
                        objetivoLista

                        ||

                        name ===
                        objetivoLista

                    );

                }

            );


        if (
            !lista
        ) {

            throw new Error(

                `No se encontró la lista "${PERMISOS_LIST_NAME}".`

            );

        }


        console.log(
            "Lista PermisosTI encontrada:",
            lista
        );


        /* -----------------------------------------
           OBTENER ITEMS
        ----------------------------------------- */

        let itemsUrl =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${sitio.id}/lists/${lista.id}/items` +

            `?$expand=fields&$top=500`;


        const items =
            [];


        while (
            itemsUrl
        ) {

            const resultado =
                await graphFetch(

                    itemsUrl,

                    token

                );


            if (

                resultado &&

                Array.isArray(
                    resultado.value
                )

            ) {

                items.push(
                    ...resultado.value
                );

            }


            itemsUrl =

                resultado[
                    "@odata.nextLink"
                ] ||

                null;

        }


        console.log(
            "Cantidad de registros PermisosTI:",
            items.length
        );


        /* -----------------------------------------
           BUSCAR REGISTRO DEL USUARIO
        ----------------------------------------- */

        const registro =

            items.find(

                item => {

                    const correoRegistro =

                        obtenerCampoFlexible(

                            item,

                            [

                                "UsuarioCorreo",

                                "Usuario Correo",

                                "Correo",

                                "Email",

                                "CorreoUsuario"

                            ]

                        );


                    const correoNormalizado =

                        normalizarCorreo(
                            correoRegistro
                        );


                    return (

                        correoNormalizado ===
                        correo

                    );

                }

            );


        /* -----------------------------------------
           SIN REGISTRO
        ----------------------------------------- */

        if (
            !registro
        ) {

            console.warn(
                "NO SE ENCONTRÓ REGISTRO PARA:",
                correo
            );


            return crearPermisosDenegados(

                correo,

                cuenta.name ||
                correo

            );

        }


        console.log(
            "REGISTRO DEL USUARIO ENCONTRADO:",
            registro
        );


        /* -----------------------------------------
           GRUPO
        ----------------------------------------- */

        const grupo =

            String(

                obtenerCampoFlexible(

                    registro,

                    [

                        "Grupo"

                    ]

                ) ||

                "Normal"

            ).trim();


        console.log(
            "Grupo del usuario:",
            grupo
        );


        /* -----------------------------------------
           GRUPO TI
           ACCESO TOTAL
        ----------------------------------------- */

        if (

            normalizarTexto(
                grupo
            ) ===

            normalizarTexto(
                "TI"
            )

        ) {

            console.log(
                "Usuario pertenece al grupo TI."
            );


            return crearPermisosTI(

                correo,

                cuenta.name ||
                correo

            );

        }


        /* -----------------------------------------
           PERMISOS INDIVIDUALES
        ----------------------------------------- */

        const permisos = {

            Oficina:
                valorBooleano(

                    obtenerCampoFlexible(

                        registro,

                        ["Oficina"]

                    )

                ),

            Personal:
                valorBooleano(

                    obtenerCampoFlexible(

                        registro,

                        ["Personal"]

                    )

                ),

            Reservas:
                valorBooleano(

                    obtenerCampoFlexible(

                        registro,

                        ["Reservas"]

                    )

                ),

            Salas:
                valorBooleano(

                    obtenerCampoFlexible(

                        registro,

                        ["Salas"]

                    )

                ),

            Comunicados:
                valorBooleano(

                    obtenerCampoFlexible(

                        registro,

                        ["Comunicados"]

                    )

                ),

            Seguridad:
                valorBooleano(

                    obtenerCampoFlexible(

                        registro,

                        ["Seguridad"]

                    )

                ),

            Infraestructura:
                valorBooleano(

                    obtenerCampoFlexible(

                        registro,

                        ["Infraestructura"]

                    )

                ),

            Tickets:
                valorBooleano(

                    obtenerCampoFlexible(

                        registro,

                        ["Tickets"]

                    )

                ),

            Configuracion:
                valorBooleano(

                    obtenerCampoFlexible(

                        registro,

                        ["Configuracion"]

                    )

                ),

            /*
             * Permisos es administrativo.
             * Solo TI puede verlo.
             */

            Permisos:
                false

        };


        console.log(
            "PERMISOS LEÍDOS DESDE SHAREPOINT:",
            permisos
        );


        return {

            correo:
                correo,

            nombre:
                cuenta.name ||

                obtenerCampoFlexible(

                    registro,

                    [

                        "NombreUsuario",

                        "Nombre Usuario",

                        "Nombre"

                    ]

                ) ||

                correo,

            grupo:
                "Normal",

            permisos:
                permisos

        };

    }


    /* =========================================================
       CREAR PERMISOS TI
    ========================================================= */

    function crearPermisosTI (

        correo,

        nombre

    ) {

        return {

            correo:
                correo,

            nombre:
                nombre,

            grupo:
                "TI",

            permisos: {

                Oficina:
                    true,

                Personal:
                    true,

                Reservas:
                    true,

                Salas:
                    true,

                Comunicados:
                    true,

                Seguridad:
                    true,

                Infraestructura:
                    true,

                Tickets:
                    true,

                Permisos:
                    true,

                Configuracion:
                    true

            }

        };

    }


    /* =========================================================
       CREAR PERMISOS DENEGADOS
    ========================================================= */

    function crearPermisosDenegados (

        correo = "",

        nombre = ""

    ) {

        return {

            correo:
                correo,

            nombre:
                nombre,

            grupo:
                "Normal",

            permisos: {

                Oficina:
                    false,

                Personal:
                    false,

                Reservas:
                    false,

                Salas:
                    false,

                Comunicados:
                    false,

                Seguridad:
                    false,

                Infraestructura:
                    false,

                Tickets:
                    false,

                Permisos:
                    false,

                Configuracion:
                    false

            }

        };

    }


    /* =========================================================
       APLICAR PERMISOS AL MENÚ
    ========================================================= */

    function aplicarPermisosMenu (

        datos

    ) {

        const permisos =

            datos &&

            datos.permisos

                ? datos.permisos

                : {};


        console.log(
            "Aplicando permisos al menú:",
            permisos
        );


        const elementos =
            document.querySelectorAll(
                ".sidebar li"
            );


        elementos.forEach(

            elemento => {

                const modulo =
                    obtenerNombreModulo(
                        elemento.textContent
                    );


                if (
                    !modulo
                ) {

                    return;

                }


                const permitido =
                    permisos[modulo] === true;


                elemento.style.display =

                    permitido

                        ? ""

                        : "none";


                console.log(

                    `Menú ${modulo}:`,

                    permitido
                        ? "PERMITIDO"
                        : "DENEGADO"

                );

            }

        );

    }


    /* =========================================================
       OBTENER MÓDULO DEL LI
    ========================================================= */

    function obtenerNombreModulo (

        texto

    ) {

        const limpio =
            normalizarTexto(
                texto
            );


        if (
            limpio.includes("oficina")
        ) {

            return "Oficina";

        }


        if (
            limpio.includes("personal")
        ) {

            return "Personal";

        }


        if (
            limpio.includes("reservas")
        ) {

            return "Reservas";

        }


        if (
            limpio.includes("salas")
        ) {

            return "Salas";

        }


        if (
            limpio.includes("comunicados")
        ) {

            return "Comunicados";

        }


        if (
            limpio.includes("seguridad")
        ) {

            return "Seguridad";

        }


        if (
            limpio.includes("infraestructura")
        ) {

            return "Infraestructura";

        }


        if (
            limpio.includes("tickets")
        ) {

            return "Tickets";

        }


        if (
            limpio.includes("permisos")
        ) {

            return "Permisos";

        }


        if (
            limpio.includes("configuracion")
        ) {

            return "Configuracion";

        }


        return null;

    }


    /* =========================================================
       VALIDAR PÁGINA ACTUAL
    ========================================================= */

    function validarPaginaActual (

        datos

    ) {

        const pagina =
            obtenerPaginaActual();


        console.log(
            "Página normalizada:",
            pagina
        );


        /*
         * Login y raíz.
         */

        if (

            !pagina ||

            pagina === "login.html"

        ) {

            return;

        }


        const modulo =
            encontrarPermisoPagina(
                pagina
            );


        console.log(
            "Módulo requerido:",
            modulo
        );


        /*
         * Página no incluida en el mapa.
         */

        if (
            !modulo
        ) {

            console.log(
                "La página no requiere permiso específico."
            );

            return;

        }


        const permitido =

            datos &&

            datos.permisos &&

            datos.permisos[modulo] === true;


        console.log(

            `Permiso para ${modulo}:`,

            permitido

        );


        if (
            permitido
        ) {

            console.log(
                `Acceso permitido a ${modulo}.`
            );

            return;

        }


        console.warn(
            `Acceso DENEGADO a ${modulo}.`
        );


        bloquearPagina(
            modulo
        );

    }


    /* =========================================================
       OBTENER PÁGINA ACTUAL
    ========================================================= */

    function obtenerPaginaActual () {

        let ruta =
            window.location.pathname;


        /*
         * Quitar /live-office/
         */

        ruta =
            ruta.replace(
                /^\/live-office\/?/i,
                ""
            );


        /*
         * Quitar / iniciales restantes.
         */

        ruta =
            ruta.replace(
                /^\/+/,
                ""
            );


        /*
         * Si estamos en:
         *
         * /live-office/
         *
         * se considera index.html.
         */

        if (
            ruta === ""
        ) {

            ruta =
                "index.html";

        }


        return ruta.toLowerCase();

    }


    /* =========================================================
       BUSCAR PERMISO DE LA PÁGINA
    ========================================================= */

    function encontrarPermisoPagina (

        pagina

    ) {

        const paginaNormalizada =
            normalizarRuta(
                pagina
            );


        const entrada =

            Object.entries(
                PAGINAS_PERMISOS
            ).find(

                ([ruta]) =>

                    normalizarRuta(
                        ruta
                    ) ===
                    paginaNormalizada

            );


        return entrada
            ? entrada[1]
            : null;

    }


    /* =========================================================
       BLOQUEAR PÁGINA
    ========================================================= */

    function bloquearPagina (

        modulo

    ) {

        const paginaActual =
            obtenerPaginaActual();


        /*
         * Si estamos en Oficina,
         * mostrar bloqueo.
         */

        if (
            paginaActual ===
            "index.html"
        ) {

            mostrarBloqueoOficina(
                modulo
            );


            return;

        }


        /*
         * Cualquier otro módulo bloqueado
         * vuelve a Oficina.
         */

        window.location.replace(
            ROOT_PATH
        );

    }


    /* =========================================================
       BLOQUEO EN OFICINA
    ========================================================= */

    function mostrarBloqueoOficina (

        modulo

    ) {

        const contenido =
            document.querySelector(
                ".content"
            );


        if (
            !contenido
        ) {

            return;

        }


        contenido.innerHTML = `

            <div

                style="

                    max-width: 600px;

                    margin: 40px auto;

                    padding: 35px;

                    background: #ffffff;

                    border-radius: 18px;

                    text-align: center;

                    box-shadow:
                        0 6px 20px
                        rgba(0,0,0,.08);

                "

            >

                <div

                    style="

                        font-size: 44px;

                        margin-bottom: 15px;

                    "

                >

                    🔒

                </div>


                <h2

                    style="

                        margin: 0 0 10px;

                        color: #0f172a;

                        font-size: 24px;

                    "

                >

                    Acceso restringido

                </h2>


                <p

                    style="

                        margin: 0;

                        color: #64748b;

                        font-size: 13px;

                        line-height: 1.5;

                    "

                >

                    No tienes permisos para acceder

                    al módulo
                    ${escaparHTML(modulo)}.

                </p>

            </div>

        `;

    }


    /* =========================================================
       OBTENER CAMPO SHAREPOINT
    ========================================================= */

    function obtenerCampo (

        registro,

        nombre

    ) {

        if (
            !registro
        ) {

            return "";

        }


        const fields =
            registro.fields || {};


        const objetivo =
            normalizarTexto(
                nombre
            );


        const clave =

            Object.keys(
                fields
            ).find(

                key =>

                    normalizarTexto(
                        key
                    ) ===
                    objetivo

            );


        if (
            !clave
        ) {

            return "";

        }


        return fields[
            clave
        ];

    }


    /* =========================================================
       OBTENER CAMPO FLEXIBLE
    ========================================================= */

    function obtenerCampoFlexible (

        registro,

        nombres

    ) {

        if (
            !registro
        ) {

            return "";

        }


        for (
            const nombre of nombres
        ) {

            const valor =
                obtenerCampo(
                    registro,
                    nombre
                );


            if (

                valor !== "" &&

                valor !== null &&

                valor !== undefined

            ) {

                return valor;

            }

        }


        return "";

    }


    /* =========================================================
       VALOR BOOLEANO
    ========================================================= */

    function valorBooleano (

        valor

    ) {

        if (

            valor === true ||

            valor === 1

        ) {

            return true;

        }


        if (

            valor === false ||

            valor === 0

        ) {

            return false;

        }


        const texto =

            String(
                valor ?? ""
            )

                .trim()

                .toLowerCase();


        return (

            texto === "true" ||

            texto === "1" ||

            texto === "yes" ||

            texto === "sí" ||

            texto === "si" ||

            texto === "checked" ||

            texto === "on"

        );

    }


    /* =========================================================
       NORMALIZAR TEXTO
    ========================================================= */

    function normalizarTexto (

        valor

    ) {

        return String(
            valor ?? ""
        )

            .normalize(
                "NFD"
            )

            .replace(
                /[\u0300-\u036f]/g,
                ""
            )

            .toLowerCase()

            .replace(
                /[^a-z0-9]/g,
                ""
            );

    }


    /* =========================================================
       NORMALIZAR CORREO
    ========================================================= */

    function normalizarCorreo (

        correo

    ) {

        return String(
            correo ?? ""
        )

            .trim()

            .toLowerCase();

    }


    /* =========================================================
       NORMALIZAR RUTA
    ========================================================= */

    function normalizarRuta (

        ruta

    ) {

        return String(
            ruta ?? ""
        )

            .replace(
                /^\/+/,
                ""
            )

            .replace(
                /\/+$/,
                ""
            )

            .toLowerCase();

    }


    /* =========================================================
       ESCAPAR HTML
    ========================================================= */

    function escaparHTML (

        valor

    ) {

        return String(
            valor ?? ""
        )

            .replace(
                /&/g,
                "&amp;"
            )

            .replace(
                /</g,
                "&lt;"
            )

            .replace(
                />/g,
                "&gt;"
            )

            .replace(
                /"/g,
                "&quot;"
            )

            .replace(
                /'/g,
                "&#039;"
            );

    }


    /* =========================================================
       DETECTAR ERROR DE SESIÓN
    ========================================================= */

    function esErrorSesion (

        error

    ) {

        if (
            !error
        ) {

            return false;

        }


        if (

            error.codigo ===
            "SIN_SESION"

        ) {

            return true;

        }


        if (

            error.codigo ===
            "DOMINIO_NO_AUTORIZADO"

        ) {

            return true;

        }


        if (

            error.status ===
            401

        ) {

            return true;

        }


        return false;

    }


    /* =========================================================
       API GLOBAL DE AUTENTICACIÓN
       
       Permite que otros módulos puedan reutilizar
       la misma instancia MSAL.
    ========================================================= */

    window.AlferzaAuth = {

        obtenerMSAL:
            crearMSAL,

        obtenerCuenta:
            obtenerCuentaActual,

        obtenerToken:
            async function (scopes) {

                const {

                    instancia,

                    cuenta

                } =
                    await obtenerCuentaActual();


                return obtenerToken(

                    instancia,

                    cuenta,

                    scopes

                );

            },

        obtenerCorreo:
            async function () {

                const {

                    cuenta

                } =
                    await obtenerCuentaActual();


                return normalizarCorreo(

                    cuenta.username ||

                    cuenta.idTokenClaims?.preferred_username ||

                    cuenta.idTokenClaims?.email ||

                    ""

                );

            }

    };

})();