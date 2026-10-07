/* =========================================
   PROTECCIÓN Y CONTROL DE PERMISOS
   ALFERZA LIVE OFFICE
========================================= */

(function protegerPagina() {

    /* =========================================
       CONFIGURACIÓN GENERAL
    ========================================= */

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


    /* =========================================
       SHAREPOINT
    ========================================= */

    const SHAREPOINT_HOST =
        "alferzaholding-my.sharepoint.com";


    const SHAREPOINT_SITE_PATH =
        "/personal/soporte1_alferza_pe";


    const PERMISOS_LIST_NAME =
        "PermisosTI";


    /* =========================================
       CUENTA TI INICIAL
       ESTA CUENTA SIEMPRE TIENE ACCESO TOTAL
    ========================================= */

    const USUARIO_TI_INICIAL =
        "soporte1@alferza.pe";


    /* =========================================
       MAPEO DE PÁGINAS
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


    if (
        sesion !==
        "true"
    ) {

        guardarReturnUrl();


        window.location.replace(
            LOGIN_PATH
        );


        return;

    }


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
       INICIO
    ========================================= */

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


    /* =========================================
       INICIAR CONTROL
    ========================================= */

    async function iniciarControlPermisos() {

        /*
         * Si ya hay permisos guardados en sesión,
         * los usamos inmediatamente para ocultar
         * el menú mientras se actualiza la
         * información desde SharePoint.
         */

        aplicarPermisosGuardados();


        try {

            const permisos =
                await obtenerPermisosUsuario();


            /*
             * Guardar permisos actuales.
             */

            sessionStorage.setItem(
                PERMISOS_KEY,
                JSON.stringify(
                    permisos
                )
            );


            /*
             * Aplicar menú.
             */

            aplicarPermisosMenu(
                permisos
            );


            /*
             * Validar página actual.
             */

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
             * En caso de error no damos
             * acceso automáticamente.
             */

            const permisosDenegados =
                crearPermisosDenegados();


            sessionStorage.setItem(
                PERMISOS_KEY,
                JSON.stringify(
                    permisosDenegados
                )
            );


            aplicarPermisosMenu(
                permisosDenegados
            );


            validarPaginaActual(
                permisosDenegados
            );

        }

    }


    /* =========================================
       APLICAR PERMISOS GUARDADOS
    ========================================= */

    function aplicarPermisosGuardados() {

        try {

            const almacenados =
                sessionStorage.getItem(
                    PERMISOS_KEY
                );


            if (!almacenados) {

                return;

            }


            const permisos =
                JSON.parse(
                    almacenados
                );


            if (
                permisos &&
                permisos.permisos
            ) {

                aplicarPermisosMenu(
                    permisos
                );

            }

        }

        catch (error) {

            console.warn(
                "No se pudieron leer permisos almacenados:",
                error
            );

        }

    }


    /* =========================================
       OBTENER MSAL
    ========================================= */

    function obtenerMSAL() {

        if (
            typeof msal !==
            "undefined"
        ) {

            return msal;

        }


        return null;

    }


    /* =========================================
       ESPERAR MSAL
       EVITA ERROR SI LA LIBRERÍA
       TODAVÍA NO TERMINÓ DE CARGAR
    ========================================= */

    async function esperarMSAL() {

        const maxIntentos =
            50;


        for (
            let intento = 0;
            intento < maxIntentos;
            intento++
        ) {

            if (
                typeof msal !==
                "undefined"
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


    /* =========================================
       CREAR MSAL
    ========================================= */

    async function crearMSAL() {

        const libreria =
            await esperarMSAL();


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
            await crearMSAL();


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


        let cuenta =
            instancia.getActiveAccount();


        if (!cuenta) {

            cuenta =
                cuentas[0];

            instancia.setActiveAccount(
                cuenta
            );

        }


        return {

            instancia:
                instancia,

            cuenta:
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


        /*
         * Obtener correo.
         */

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


        /*
         * =====================================
         * TI INICIAL
         * =====================================
         */

        if (
            correo ===
            normalizarCorreo(
                USUARIO_TI_INICIAL
            )
        ) {

            return crearPermisosTI(

                correo,

                cuenta.name ||
                correo

            );

        }


        /*
         * =====================================
         * SITIO SHAREPOINT
         * =====================================
         */

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


        /*
         * =====================================
         * OBTENER LISTAS
         * =====================================
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


        const objetivoLista =
            normalizarTexto(
                PERMISOS_LIST_NAME
            );


        const lista =
            (
                listas.value ||
                []
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


        if (!lista) {

            throw new Error(
                `No se encontró la lista "${PERMISOS_LIST_NAME}".`
            );

        }


        /*
         * =====================================
         * OBTENER ITEMS DE PERMISOSTI
         * =====================================
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


        /*
         * =====================================
         * BUSCAR REGISTRO DEL USUARIO
         * =====================================
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
                        ) ===
                        correo

                    );

                }

            );


        /*
         * =====================================
         * SIN REGISTRO
         * =====================================
         */

        if (!registro) {

            return crearPermisosDenegados(

                correo,

                cuenta.name ||
                correo

            );

        }


        /*
         * =====================================
         * GRUPO
         * =====================================
         */

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
         * =====================================
         * GRUPO TI
         * ACCESO TOTAL
         * =====================================
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
         * =====================================
         * GRUPO NORMAL
         * =====================================
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

                /*
                 * La página Permisos es administrativa
                 * y no tiene columna propia en PermisosTI.
                 * Solo Grupo TI puede verla.
                 */

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
       USA EXACTAMENTE TU SIDEBAR:
       
       <aside class="sidebar">
           <h2>MENÚ</h2>
           <ul>
               <li>...</li>
           </ul>
       </aside>
    ========================================= */

    function aplicarPermisosMenu(
        datos
    ) {

        const permisos =
            datos &&
            datos.permisos
                ? datos.permisos
                : {};


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
       OBTENER MÓDULO DEL LI
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
         * Login y raíz no requieren
         * permiso de módulo.
         */

        if (
            !pagina ||
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
         * Archivo no incluido en el mapa:
         * no aplicar restricción.
         */

        if (!modulo) {

            return;

        }


        const permitido =
            datos &&
            datos.permisos &&
            datos.permisos[
                modulo
            ] === true;


        if (
            permitido
        ) {

            return;

        }


        /*
         * Si no tiene permiso para la página
         * actual, no se muestra.
         */

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


        return ruta
            .toLowerCase();

    }


    /* =========================================
       BUSCAR PERMISO DE LA PÁGINA
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

        const paginaActual =
            obtenerPaginaActual();


        /*
         * Si ya estamos en Oficina y no hay
         * acceso, mostramos bloqueo para evitar
         * un bucle infinito.
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
            "/live-office/index.html"
        );

    }


    /* =========================================
       BLOQUEO EN LA PÁGINA OFICINA
    ========================================= */

    function mostrarBloqueoOficina(
        modulo
    ) {

        const contenido =
            document.querySelector(
                ".content"
            );


        if (!contenido) {

            return;

        }


        contenido.innerHTML =

            `
            <div
                style="
                    max-width: 600px;
                    margin: 40px auto;
                    padding: 35px;
                    background: #ffffff;
                    border-radius: 18px;
                    text-align: center;
                    box-shadow: 0 6px 20px rgba(0,0,0,.08);
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
                    al módulo ${escaparHTML(modulo)}.
                </p>

            </div>
            `;

    }


    /* =========================================
       OBTENER CAMPO SHAREPOINT
    ========================================= */

    function obtenerCampo(
        registro,
        nombre
    ) {

        if (
            !registro
        ) {

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