/* =========================================
   PROTECCIÓN DE ACCESO
   ALFERZA LIVE OFFICE
========================================= */

(function protegerPagina() {

    const SESION_KEY = "alferza_login";
    const RETURN_KEY = "alferza_return_url";
    const PERMISOS_KEY = "alferza_permisos";

    const LOGIN_PATH = "/live-office/login.html";
    const ROOT_PATH = "/live-office/";

    const SHAREPOINT_HOST =
        "alferzaholding-my.sharepoint.com";

    const SHAREPOINT_SITE_PATH =
        "/personal/soporte1_alferza_pe";

    const PERMISOS_LIST_NAME =
        "PermisosTI";

    const USUARIO_TI_INICIAL =
        "soporte1@alferza.pe";


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
       COMPROBAR SESIÓN
    ========================================= */

    const sesion =
        sessionStorage.getItem(
            SESION_KEY
        );


    if (sesion !== "true") {

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

            const permisos =
                await obtenerPermisosUsuario();


            sessionStorage.setItem(
                PERMISOS_KEY,
                JSON.stringify(permisos)
            );


            aplicarPermisosMenu(
                permisos
            );


            validarPaginaActual(
                permisos
            );

        }

        catch (error) {

            console.error(
                "Error comprobando permisos:",
                error
            );


            /*
             * Si no se pudo consultar SharePoint,
             * no mostramos automáticamente todos
             * los módulos.
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
       OBTENER MSAL
    ========================================= */

    function obtenerMSAL() {

        if (
            typeof msal !== "undefined"
        ) {

            return msal;

        }

        return null;

    }


    /* =========================================
       OBTENER INSTANCIA MSAL
    ========================================= */

    function crearMSAL() {

        const libreria =
            obtenerMSAL();


        if (!libreria) {

            throw new Error(
                "MSAL no está disponible."
            );

        }


        const config = {

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


        return new libreria.PublicClientApplication(
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
                "No existe una cuenta de Microsoft."
            );

        }


        const cuenta =
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

        const resultado =
            await instancia.acquireTokenSilent({

                scopes: [
                    "User.Read",
                    "Sites.ReadWrite.All"
                ],

                account:
                    cuenta

            });


        return resultado.accessToken;

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
                ""

            );


        if (!correo) {

            throw new Error(
                "No se pudo determinar el correo del usuario."
            );

        }


        /*
         * TI inicial:
         * acceso completo.
         */

        if (
            correo ===
            USUARIO_TI_INICIAL
        ) {

            return crearPermisosTI(
                correo,
                cuenta.name ||
                correo
            );

        }


        /*
         * Obtener sitio SharePoint.
         */

        const sitioUrl =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;


        const sitio =
            await graphFetch(
                sitioUrl,
                token
            );


        if (!sitio?.id) {

            throw new Error(
                "No se pudo obtener el sitio de SharePoint."
            );

        }


        /*
         * Obtener listas.
         */

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
                listas.value || []
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


        /*
         * Obtener elementos de PermisosTI.
         */

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


        /*
         * Buscar usuario por correo.
         */

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


        /*
         * Si no está registrado:
         * no tiene acceso.
         */

        if (!registro) {

            return crearPermisosDenegados(
                correo,
                cuenta.name ||
                correo
            );

        }


        const grupo =
            String(

                obtenerCampo(
                    registro,
                    "Grupo"
                ) ||

                "Normal"

            )
                .trim();


        /*
         * Grupo TI:
         * acceso total.
         */

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


        /*
         * Grupo Normal:
         * usar exactamente los valores
         * guardados en PermisosTI.
         */

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

            permisos: {

                Oficina:
                    valorBooleano(
                        obtenerCampo(
                            registro,
                            "Oficina"
                        )
                    ),

                Personal:
                    valorBooleano(
                        obtenerCampo(
                            registro,
                            "Personal"
                        )
                    ),

                Reservas:
                    valorBooleano(
                        obtenerCampo(
                            registro,
                            "Reservas"
                        )
                    ),

                Salas:
                    valorBooleano(
                        obtenerCampo(
                            registro,
                            "Salas"
                        )
                    ),

                Comunicados:
                    valorBooleano(
                        obtenerCampo(
                            registro,
                            "Comunicados"
                        )
                    ),

                Seguridad:
                    valorBooleano(
                        obtenerCampo(
                            registro,
                            "Seguridad"
                        )
                    ),

                Infraestructura:
                    valorBooleano(
                        obtenerCampo(
                            registro,
                            "Infraestructura"
                        )
                    ),

                Tickets:
                    valorBooleano(
                        obtenerCampo(
                            registro,
                            "Tickets"
                        )
                    ),

                Configuracion:
                    valorBooleano(
                        obtenerCampo(
                            registro,
                            "Configuracion"
                        )
                    ),

                Permisos:
                    false

            }

        };

    }


    /* =========================================
       CREAR PERMISOS TI
    ========================================= */

    function crearPermisosTI(
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


    /* =========================================
       CREAR PERMISOS DENEGADOS
    ========================================= */

    function crearPermisosDenegados(
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

                const texto =
                    obtenerNombreModulo(
                        elemento.textContent
                    );


                if (!texto) {

                    return;

                }


                const permitido =
                    permisos[
                        texto
                    ] === true;


                /*
                 * TI:
                 * todos visibles.
                 *
                 * Normal:
                 * solo permitidos.
                 */

                elemento.style.display =
                    permitido
                        ? ""
                        : "none";

            }
        );

    }


    /* =========================================
       OBTENER NOMBRE DEL MÓDULO DEL LI
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
         * Login y raíz:
         * no necesitan permisos de módulo.
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


        /*
         * Página no registrada:
         * dejar continuar.
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


        ruta =
            ruta.replace(
                /^\/live-office\//,
                ""
            );


        ruta =
            ruta.replace(
                /^\/+/,
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


        /*
         * Evitamos mostrar el contenido
         * restringido.
         */

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

            font-family: 'Segoe UI', Arial, sans-serif;

            text-align: center;

        `;


        mensaje.innerHTML = `

            <div style="
                max-width: 500px;
                padding: 40px;
                background: white;
                border-radius: 18px;
                box-shadow: 0 6px 20px rgba(0,0,0,.08);
            ">

                <div style="
                    font-size: 44px;
                    margin-bottom: 18px;
                ">
                    🔒
                </div>

                <h2 style="
                    margin: 0 0 10px;
                    font-size: 24px;
                ">
                    Acceso restringido
                </h2>

                <p style="
                    margin: 0 0 24px;
                    color: #64748b;
                    font-size: 13px;
                    line-height: 1.5;
                ">
                    No tienes permisos para acceder a
                    <strong>${escaparHTML(modulo)}</strong>.
                </p>

                <button
                    type="button"
                    id="volverInicioPermisos"
                    style="
                        padding: 11px 18px;
                        border: none;
                        border-radius: 10px;
                        background: #0a84ff;
                        color: white;
                        font-family: inherit;
                        font-size: 12px;
                        font-weight: 800;
                        cursor: pointer;
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
                        "/live-office/index.html";

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