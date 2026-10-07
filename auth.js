/* =========================================
   PROTECCIÓN DE ACCESO
   ALFERZA LIVE OFFICE
========================================= */

(function protegerPagina() {

    /* =========================================
       CONFIGURACIÓN
    ========================================= */

    const SESION_KEY = "alferza_login";
    const RETURN_KEY = "alferza_return_url";
    const PERMISOS_KEY = "alferza_permisos";

    const ROOT_PATH = "/live-office/";
    const LOGIN_PATH = "/live-office/login.html";
    const BLANK_PATH = "/live-office/blank.html";

    const CLIENT_ID =
        "5d98417c-74a7-4fab-8f2c-41ac127be696";

    const AUTHORITY =
        "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1";

    const REDIRECT_URI =
        window.location.origin + BLANK_PATH;

    const SHAREPOINT_HOST =
        "alferzaholding-my.sharepoint.com";

    const SHAREPOINT_SITE_PATH =
        "/personal/soporte1_alferza_pe";

    const PERMISOS_LIST_NAME =
        "PermisosTI";

    const USUARIO_TI_INICIAL =
        "soporte1@alferza.pe";


    /* =========================================
       PERMISOS NECESARIOS PARA LEER
       PermisosTI
    ========================================= */

    const SCOPES = [
        "User.Read",
        "Sites.Read.All"
    ];


    /* =========================================
       MAPEO DE PÁGINAS → PERMISOS
    ========================================= */

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


    /* =========================================
       PERMISOS DISPONIBLES
    ========================================= */

    const MODULOS = [

        "Oficina",
        "Personal",
        "Reservas",
        "Salas",
        "Comunicados",
        "Seguridad",
        "Infraestructura",
        "Tickets",
        "Permisos",
        "Configuracion"

    ];


    /* =========================================
       COMPROBAR SESIÓN LOCAL
    ========================================= */

    const sesion =
        sessionStorage.getItem(
            SESION_KEY
        );


    /*
     * login.html y blank.html no necesitan
     * esta protección.
     */

    const rutaActual =
        window.location.pathname;


    if (
        rutaActual === LOGIN_PATH ||
        rutaActual === BLANK_PATH
    ) {

        return;

    }


    /*
     * Si no existe la sesión creada por login.js,
     * mandamos al login.
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


    /* =========================================
       INICIAR PROTECCIÓN
    ========================================= */

    iniciarControlPermisos();


    /* =========================================
       GUARDAR URL DE RETORNO
    ========================================= */

    function guardarReturnUrl() {

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


    /* =========================================
       INICIAR CONTROL
    ========================================= */

    async function iniciarControlPermisos() {

        try {

            /*
             * Esperamos al DOM si fuera necesario.
             */

            if (
                document.readyState ===
                "loading"
            ) {

                await new Promise(
                    resolve => {

                        document.addEventListener(
                            "DOMContentLoaded",
                            resolve,
                            {
                                once: true
                            }
                        );

                    }
                );

            }


            console.log(
                "ALFERZA LIVE OFFICE: verificando permisos..."
            );


            const datos =
                await obtenerPermisosUsuario();


            console.log(
                "Permisos obtenidos:",
                datos
            );


            /*
             * Guardamos únicamente permisos
             * obtenidos correctamente.
             */

            sessionStorage.setItem(

                PERMISOS_KEY,

                JSON.stringify(
                    datos
                )

            );


            /*
             * Aplicar menú.
             */

            aplicarPermisosMenu(
                datos
            );


            /*
             * Validar página actual.
             */

            validarPaginaActual(
                datos
            );


        }

        catch (error) {

            console.error(
                "Error comprobando permisos:",
                error
            );


            /*
             * IMPORTANTE:
             *
             * Si SharePoint falla NO damos acceso.
             * Tampoco intentamos iniciar sesión otra vez.
             */

            const permisosBloqueados =
                crearPermisosDenegados();


            sessionStorage.setItem(

                PERMISOS_KEY,

                JSON.stringify(
                    permisosBloqueados
                )

            );


            aplicarPermisosMenu(
                permisosBloqueados
            );


            validarPaginaActual(
                permisosBloqueados
            );

        }

    }


    /* =========================================
       CREAR MSAL
    ========================================= */

    function crearMSAL() {

        if (
            typeof msal === "undefined"
        ) {

            throw new Error(
                "MSAL no está disponible."
            );

        }


        const config = {

            auth: {

                clientId:
                    CLIENT_ID,

                authority:
                    AUTHORITY,

                /*
                 * SIEMPRE el blank de la raíz.
                 */

                redirectUri:
                    REDIRECT_URI

            },

            cache: {

                cacheLocation:
                    "sessionStorage",

                storeAuthStateInCookie:
                    false

            }

        };


        return new msal.PublicClientApplication(
            config
        );

    }


    /* =========================================
       OBTENER CUENTA ACTUAL
    ========================================= */

    async function obtenerCuentaActual() {

        const instancia =
            crearMSAL();


        const cuentas =
            instancia.getAllAccounts();


        if (
            !cuentas ||
            !cuentas.length
        ) {

            throw new Error(
                "No existe una cuenta de Microsoft en la sesión."
            );

        }


        /*
         * Intentamos encontrar la cuenta
         * correspondiente a ALFERZA.
         */

        const cuentaAlferza =
            cuentas.find(

                cuenta => {

                    const correo =
                        normalizarCorreo(

                            cuenta.username ||

                            cuenta.idTokenClaims?.preferred_username ||

                            ""

                        );


                    return correo.endsWith(
                        "@alferza.pe"
                    );

                }

            );


        const cuenta =
            cuentaAlferza ||
            cuentas[0];


        instancia.setActiveAccount(
            cuenta
        );


        return {

            instancia,
            cuenta

        };

    }


    /* =========================================
       OBTENER TOKEN
    ========================================= */

    async function obtenerToken(
        instancia,
        cuenta
    ) {

        try {

            const resultado =
                await instancia.acquireTokenSilent({

                    scopes:
                        SCOPES,

                    account:
                        cuenta

                });


            return resultado.accessToken;

        }

        catch (error) {

            console.error(
                "Error obteniendo token silencioso:",
                error
            );


            /*
             * NO hacemos loginPopup.
             *
             * login.js es el único responsable
             * de iniciar sesión.
             */

            throw new Error(
                "No se pudo obtener el token de Microsoft para consultar PermisosTI."
            );

        }

    }


    /* =========================================
       GRAPH FETCH
    ========================================= */

    async function graphFetch(
        url,
        token
    ) {

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


            throw new Error(

                `Graph ${respuesta.status}: ${texto}`

            );

        }


        return respuesta.json();

    }


    /* =========================================
       OBTENER PERMISOS DEL USUARIO
    ========================================= */

    async function obtenerPermisosUsuario() {

        const {

            instancia,
            cuenta

        } =
            await obtenerCuentaActual();


        const token =
            await obtenerToken(

                instancia,

                cuenta

            );


        const correo =
            normalizarCorreo(

                cuenta.username ||

                cuenta.idTokenClaims?.preferred_username ||

                cuenta.idTokenClaims?.email ||

                ""

            );


        if (!correo) {

            throw new Error(
                "No se pudo determinar el correo del usuario."
            );

        }


        console.log(
            "Usuario actual:",
            correo
        );


        /* =====================================
           TI INICIAL
        ===================================== */

        if (
            correo ===
            USUARIO_TI_INICIAL
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


        /* =====================================
           OBTENER SITIO SHAREPOINT
        ===================================== */

        const sitioUrl =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;


        const sitio =
            await graphFetch(

                sitioUrl,

                token

            );


        if (
            !sitio ||
            !sitio.id
        ) {

            throw new Error(
                "No se pudo obtener el sitio de SharePoint."
            );

        }


        console.log(
            "Sitio SharePoint:",
            sitio.id
        );


        /* =====================================
           OBTENER LISTAS
        ===================================== */

        const listasUrl =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${sitio.id}/lists` +

            `?$select=id,name,displayName&$top=200`;


        const listas =
            await graphFetch(

                listasUrl,

                token

            );


        const objetivo =
            normalizarTexto(

                PERMISOS_LIST_NAME

            );


        const lista =
            (
                listas.value ||
                []
            ).find(

                item =>

                    normalizarTexto(
                        item.displayName
                    ) === objetivo

                    ||

                    normalizarTexto(
                        item.name
                    ) === objetivo

            );


        if (!lista) {

            throw new Error(

                `No se encontró la lista "${PERMISOS_LIST_NAME}".`

            );

        }


        console.log(
            "Lista PermisosTI encontrada:",
            lista
        );


        /* =====================================
           OBTENER ITEMS
        ===================================== */

        let itemsUrl =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${sitio.id}/lists/${lista.id}/items` +

            `?$expand=fields&$top=500`;


        const items = [];


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
            "Registros PermisosTI:",
            items.length
        );


        /* =====================================
           BUSCAR USUARIO
        ===================================== */

        const registro =
            items.find(

                item => {

                    const correoRegistro =
                        obtenerCampo(

                            item,

                            "UsuarioCorreo"

                        );


                    return (

                        normalizarCorreo(
                            correoRegistro
                        ) === correo

                    );

                }

            );


        /* =====================================
           USUARIO NO REGISTRADO
        ===================================== */

        if (!registro) {

            console.warn(
                "Usuario no encontrado en PermisosTI:",
                correo
            );


            return crearPermisosDenegados(

                correo,

                cuenta.name ||
                correo

            );

        }


        console.log(
            "Registro del usuario:",
            registro
        );


        /* =====================================
           GRUPO
        ===================================== */

        const grupo =

            String(

                obtenerCampo(

                    registro,

                    "Grupo"

                ) ||

                "Normal"

            )
                .trim();


        /* =====================================
           GRUPO TI
        ===================================== */

        if (

            normalizarTexto(
                grupo
            ) ===

            normalizarTexto(
                "TI"
            )

        ) {

            return crearPermisosTI(

                correo,

                cuenta.name ||
                correo

            );

        }


        /* =====================================
           GRUPO NORMAL
        ===================================== */

        const permisos = {};


        MODULOS.forEach(

            modulo => {

                permisos[
                    modulo
                ] = false;

            }

        );


        /*
         * Permisos normales:
         * leer exactamente los valores
         * de PermisosTI.
         */

        MODULOS.forEach(

            modulo => {

                permisos[
                    modulo
                ] =

                    valorBooleano(

                        obtenerCampo(

                            registro,

                            modulo

                        )

                    );

            }

        );


        console.log(
            "Permisos del usuario:",
            permisos
        );


        return {

            correo:

                correo,

            nombre:

                cuenta.name ||

                obtenerCampo(

                    registro,

                    "NombreUsuario"

                ) ||

                correo,

            grupo:

                "Normal",

            permisos:

                permisos

        };

    }


    /* =========================================
       CREAR PERMISOS TI
    ========================================= */

    function crearPermisosTI(
        correo,
        nombre
    ) {

        const permisos = {};


        MODULOS.forEach(

            modulo => {

                permisos[
                    modulo
                ] = true;

            }

        );


        return {

            correo:
                correo,

            nombre:
                nombre,

            grupo:
                "TI",

            permisos:
                permisos

        };

    }


    /* =========================================
       CREAR PERMISOS DENEGADOS
    ========================================= */

    function crearPermisosDenegados(
        correo = "",
        nombre = ""
    ) {

        const permisos = {};


        MODULOS.forEach(

            modulo => {

                permisos[
                    modulo
                ] = false;

            }

        );


        return {

            correo:
                correo,

            nombre:
                nombre,

            grupo:
                "Normal",

            permisos:
                permisos

        };

    }


    /* =========================================
       APLICAR PERMISOS AL MENÚ
    ========================================= */

    function aplicarPermisosMenu(
        datos
    ) {

        const permisos =
            datos?.permisos || {};


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


                if (!modulo) {

                    return;

                }


                const permitido =
                    permisos[
                        modulo
                    ] === true;


                elemento.style.display =
                    permitido
                        ? ""
                        : "none";

            }

        );

    }


    /* =========================================
       OBTENER NOMBRE DEL MÓDULO
    ========================================= */

    function obtenerNombreModulo(
        texto
    ) {

        const limpio =
            normalizarTexto(
                texto
            );


        if (
            limpio.includes(
                "oficina"
            )
        ) {

            return "Oficina";

        }


        if (
            limpio.includes(
                "personal"
            )
        ) {

            return "Personal";

        }


        if (
            limpio.includes(
                "reservas"
            )
        ) {

            return "Reservas";

        }


        if (
            limpio.includes(
                "salas"
            )
        ) {

            return "Salas";

        }


        if (
            limpio.includes(
                "comunicados"
            )
        ) {

            return "Comunicados";

        }


        if (
            limpio.includes(
                "seguridad"
            )
        ) {

            return "Seguridad";

        }


        if (
            limpio.includes(
                "infraestructura"
            )
        ) {

            return "Infraestructura";

        }


        if (
            limpio.includes(
                "tickets"
            )
        ) {

            return "Tickets";

        }


        if (
            limpio.includes(
                "permisos"
            )
        ) {

            return "Permisos";

        }


        if (
            limpio.includes(
                "configuracion"
            )
        ) {

            return "Configuracion";

        }


        return null;

    }


    /* =========================================
       VALIDAR PÁGINA ACTUAL
    ========================================= */

    function validarPaginaActual(
        datos
    ) {

        const pagina =
            obtenerPaginaActual();


        /*
         * La raíz no tiene módulo específico.
         */

        if (
            !pagina
        ) {

            return;

        }


        if (
            pagina ===
            "login.html"
        ) {

            return;

        }


        const modulo =
            encontrarPermisoPagina(

                pagina

            );


        /*
         * Página no registrada.
         */

        if (!modulo) {

            return;

        }


        const permitido =
            datos?.permisos?.[

                modulo

            ] === true;


        if (
            permitido
        ) {

            console.log(
                `Acceso permitido: ${modulo}`
            );

            return;

        }


        bloquearPagina(
            modulo
        );

    }


    /* =========================================
       OBTENER PÁGINA ACTUAL
    ========================================= */

    function obtenerPaginaActual() {

        let ruta =
            window.location.pathname;


        /*
         * Quitar /live-office/
         */

        const prefijo =
            "/live-office/";


        if (
            ruta.startsWith(
                prefijo
            )
        ) {

            ruta =
                ruta.substring(
                    prefijo.length
                );

        }


        /*
         * Quitar slash inicial/final.
         */

        ruta =
            ruta.replace(
                /^\/+/,
                ""
            );


        ruta =
            ruta.replace(
                /\/+$/,
                ""
            );


        return ruta;

    }


    /* =========================================
       BUSCAR PERMISO DE PÁGINA
    ========================================= */

    function encontrarPermisoPagina(
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


    /* =========================================
       BLOQUEAR PÁGINA
    ========================================= */

    function bloquearPagina(
        modulo
    ) {

        console.warn(
            `Acceso denegado al módulo: ${modulo}`
        );


        document.body.innerHTML = "";


        const mensaje =
            document.createElement(
                "div"
            );


        mensaje.style.cssText = `

            min-height: 100vh;

            display: flex;

            align-items: center;

            justify-content: center;

            padding: 30px;

            box-sizing: border-box;

            background: #eef2f7;

            color: #0f172a;

            font-family:
                'Segoe UI',
                Arial,
                sans-serif;

            text-align: center;

        `;


        mensaje.innerHTML = `

            <div style="

                max-width: 500px;

                padding: 40px;

                background: white;

                border-radius: 18px;

                box-shadow:
                    0 6px 20px
                    rgba(0,0,0,.08);

            ">

                <div style="

                    font-size: 44px;

                    margin-bottom: 18px;

                ">

                    🔒

                </div>


                <h2 style="

                    margin:
                        0 0 10px;

                    font-size:
                        24px;

                ">

                    Acceso restringido

                </h2>


                <p style="

                    margin:
                        0 0 24px;

                    color:
                        #64748b;

                    font-size:
                        13px;

                    line-height:
                        1.5;

                ">

                    No tienes permisos para acceder a

                    <strong>
                        ${escaparHTML(modulo)}
                    </strong>.

                </p>


                <button

                    type="button"

                    id="volverInicioPermisos"

                    style="

                        padding:
                            11px 18px;

                        border:
                            none;

                        border-radius:
                            10px;

                        background:
                            #0a84ff;

                        color:
                            white;

                        font-family:
                            inherit;

                        font-size:
                            12px;

                        font-weight:
                            800;

                        cursor:
                            pointer;

                    "

                >

                    Volver a Oficina

                </button>

            </div>

        `;


        document.body.appendChild(
            mensaje
        );


        const boton =
            document.getElementById(
                "volverInicioPermisos"
            );


        if (boton) {

            boton.addEventListener(

                "click",

                function () {

                    window.location.href =
                        ROOT_PATH +
                        "index.html";

                }

            );

        }

    }


    /* =========================================
       OBTENER CAMPO SHAREPOINT
    ========================================= */

    function obtenerCampo(
        registro,
        nombre
    ) {

        if (!registro) {

            return "";

        }


        const fields =
            registro.fields ||
            {};


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
                    ) === objetivo

            );


        if (!clave) {

            return "";

        }


        return fields[
            clave
        ];

    }


    /* =========================================
       VALOR BOOLEANO
    ========================================= */

    function valorBooleano(
        valor
    ) {

        if (
            valor === true ||
            valor === 1
        ) {

            return true;

        }


        const texto =
            String(
                valor || ""
            )
                .trim()
                .toLowerCase();


        return (

            texto === "true" ||

            texto === "1" ||

            texto === "yes" ||

            texto === "sí" ||

            texto === "si"

        );

    }


    /* =========================================
       NORMALIZAR TEXTO
    ========================================= */

    function normalizarTexto(
        valor
    ) {

        return String(
            valor || ""
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


    /* =========================================
       NORMALIZAR CORREO
    ========================================= */

    function normalizarCorreo(
        correo
    ) {

        return String(
            correo || ""
        )
            .trim()
            .toLowerCase();

    }


    /* =========================================
       NORMALIZAR RUTA
    ========================================= */

    function normalizarRuta(
        ruta
    ) {

        return String(
            ruta || ""
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


    /* =========================================
       ESCAPAR HTML
    ========================================= */

    function escaparHTML(
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


})();