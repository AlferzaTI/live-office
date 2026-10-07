/* =========================================================
   PROTECCIÓN Y CONTROL DE PERMISOS
   ALFERZA LIVE OFFICE
========================================================= */

(function protegerPagina () {

    /* =========================================================
       CONFIGURACIÓN GENERAL
    ========================================================= */

    const SESION_KEY = "alferza_login";
    const RETURN_KEY = "alferza_return_url";
    const PERMISOS_KEY = "alferza_permisos";

    const LOGIN_PATH = "/live-office/login.html";
    const ROOT_PATH = "/live-office/";

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

       IMPORTANTE:
       SIEMPRE USAR EL BLANK.HTML DE LA RAÍZ.
       NO CALCULARLO RELATIVO A LA PÁGINA ACTUAL.
    ========================================================= */

    const MSAL_CONFIG = {

        auth: {

            clientId:
                "5d98417c-74a7-4fab-8f2c-41ac127be696",

            authority:
                "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",

            redirectUri:
                window.location.origin +
                "/live-office/blank.html"
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

    if (sesion !== "true") {

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
            paginaActual !== ROOT_PATH &&
            !paginaActual.endsWith("/login.html")
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
        document.readyState === "loading"
    ) {

        document.addEventListener(
            "DOMContentLoaded",
            iniciarControlPermisos
        );

    } else {

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

        try {

            const permisos =
                await obtenerPermisosUsuario();

            console.log(
                "========================================="
            );

            console.log(
                "PERMISOS OBTENIDOS:",
                permisos
            );

            console.log(
                "========================================="
            );

            /*
             * Guardamos los permisos junto con el correo.
             * Así nunca reutilizamos permisos de otro usuario.
             */

            sessionStorage.setItem(
                PERMISOS_KEY,
                JSON.stringify({
                    correo: permisos.correo || "",
                    permisos: permisos.permisos || {},
                    grupo: permisos.grupo || "Normal",
                    nombre: permisos.nombre || ""
                })
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
             * Si la sesión de Microsoft ya no existe,
             * limpiamos completamente la sesión local.
             */

            if (
                esErrorSesion(error)
            ) {

                limpiarSesion();

                guardarReturnUrl();

                window.location.replace(
                    LOGIN_PATH
                );

                return;
            }

            /*
             * IMPORTANTE:
             *
             * Si SharePoint falla NO reutilizamos permisos
             * antiguos.
             *
             * Esto evita accesos incorrectos.
             */

            console.warn(
                "No se pudieron obtener los permisos actuales desde SharePoint."
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
       LIMPIAR SESIÓN
    ========================================================= */

    function limpiarSesion () {

        sessionStorage.removeItem(
            SESION_KEY
        );

        sessionStorage.removeItem(
            PERMISOS_KEY
        );

        sessionStorage.removeItem(
            RETURN_KEY
        );
    }

    /* =========================================================
       ESPERAR MSAL
    ========================================================= */

    async function esperarMSAL () {

        const maxIntentos = 100;

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

        const error =
            new Error(
                "MSAL no está disponible."
            );

        error.codigo =
            "MSAL_NO_DISPONIBLE";

        throw error;
    }

    /* =========================================================
       CREAR MSAL
    ========================================================= */

    let instanciaMSAL = null;

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

        console.log(
            "CUENTAS MSAL ENCONTRADAS:",
            cuentas
        );

        if (
            !cuentas ||
            cuentas.length === 0
        ) {

            const error =
                new Error(
                    "No existe una cuenta de Microsoft activa."
                );

            error.codigo =
                "SIN_SESION";

            throw error;
        }

        let cuenta =
            instancia.getActiveAccount();

        /*
         * Si no existe cuenta activa,
         * usamos la primera cuenta almacenada.
         */

        if (!cuenta) {

            cuenta =
                cuentas[0];

            instancia.setActiveAccount(
                cuenta
            );
        }

        console.log(
            "CUENTA MICROSOFT ACTIVA:",
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

            if (
                !resultado ||
                !resultado.accessToken
            ) {

                const error =
                    new Error(
                        "Microsoft no devolvió un token válido."
                    );

                error.codigo =
                    "TOKEN_INVALIDO";

                throw error;
            }

            return resultado.accessToken;

        }
        catch (error) {

            console.error(
                "ERROR acquireTokenSilent:",
                error
            );

            /*
             * NO hacemos loginRedirect aquí.
             *
             * Si auth.js hiciera una redirección automática,
             * podría provocar un bucle de login.
             */

            const nuevoError =
                new Error(
                    "No se pudo obtener el token de Microsoft."
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

    /* =========================================================
       OBTENER PERMISOS DEL USUARIO
    ========================================================= */

    async function obtenerPermisosUsuario () {

        const {
            instancia,
            cuenta
        } =
            await obtenerCuentaActual();

        /* =====================================================
           OBTENER CORREO
        ===================================================== */

        const correo =
            normalizarCorreo(
                cuenta.username ||
                cuenta.idTokenClaims?.preferred_username ||
                cuenta.idTokenClaims?.email ||
                cuenta.idTokenClaims?.upn ||
                ""
            );

        console.log(
            "========================================="
        );

        console.log(
            "CORREO DETECTADO:",
            correo
        );

        console.log(
            "NOMBRE CUENTA:",
            cuenta.name
        );

        console.log(
            "========================================="
        );

        if (!correo) {

            const error =
                new Error(
                    "No se pudo determinar el correo del usuario."
                );

            error.codigo =
                "CORREO_NO_DETECTADO";

            throw error;
        }

        /* =====================================================
           VALIDAR DOMINIO
        ===================================================== */

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

        /* =====================================================
           CUENTA TI INICIAL
        ===================================================== */

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
                cuenta.name || correo
            );
        }

        /* =====================================================
           TOKEN
        ===================================================== */

        const token =
            await obtenerToken(
                instancia,
                cuenta
            );

        /* =====================================================
           SITIO SHAREPOINT
        ===================================================== */

        const sitioUrl =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;

        const sitio =
            await graphFetch(
                sitioUrl,
                token
            );

        console.log(
            "SITIO SHAREPOINT:",
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

        /* =====================================================
           OBTENER LISTAS
        ===================================================== */

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
            "LISTAS SHAREPOINT:",
            listas
        );

        const objetivoLista =
            normalizarTexto(
                PERMISOS_LIST_NAME
            );

        const lista =
            (
                listas?.value || []
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
                        displayName === objetivoLista ||
                        name === objetivoLista
                    );
                }
            );

        if (!lista) {

            throw new Error(
                `No se encontró la lista "${PERMISOS_LIST_NAME}".`
            );
        }

        console.log(
            "LISTA PERMISOSTI ENCONTRADA:",
            lista
        );

        /* =====================================================
           OBTENER COLUMNAS DE LA LISTA

           Esto permite detectar correctamente los nombres
           internos de SharePoint.
        ===================================================== */

        const columnasUrl =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists/${lista.id}/columns` +
            `?$select=id,name,displayName`;

        const columnas =
            await graphFetch(
                columnasUrl,
                token
            );

        console.log(
            "COLUMNAS PERMISOSTI:",
            columnas
        );

        /* =====================================================
           MAPEAR COLUMNAS
        ===================================================== */

        const mapaColumnas =
            construirMapaColumnas(
                columnas?.value || []
            );

        console.log(
            "MAPA DE COLUMNAS:",
            mapaColumnas
        );

        /* =====================================================
           OBTENER ITEMS
        ===================================================== */

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
                resultado &&
                resultado["@odata.nextLink"]
                    ? resultado["@odata.nextLink"]
                    : null;
        }

        console.log(
            "CANTIDAD DE REGISTROS PERMISOSTI:",
            items.length
        );

        /* =====================================================
           MOSTRAR REGISTROS
        ===================================================== */

        console.log(
            "========================================="
        );

        console.log(
            "REGISTROS DE PERMISOSTI:"
        );

        items.forEach(
            item => {

                console.log(
                    "ID:",
                    item.id,

                    "| UsuarioCorreo:",
                    obtenerCampoFlexible(
                        item,
                        [
                            "UsuarioCorreo",
                            "Usuario Correo",
                            "Correo",
                            "Email",
                            "CorreoUsuario",
                            "Title"
                        ],
                        mapaColumnas
                    ),

                    "| Nombre:",
                    obtenerCampoFlexible(
                        item,
                        [
                            "NombreUsuario",
                            "Nombre Usuario",
                            "Nombre"
                        ],
                        mapaColumnas
                    )
                );
            }
        );

        console.log(
            "========================================="
        );

        /* =====================================================
           BUSCAR REGISTRO DEL USUARIO
        ===================================================== */

        let registro =
            null;

        for (
            const item of items
        ) {

            const posiblesCorreos = [

                obtenerCampoFlexible(
                    item,
                    [
                        "UsuarioCorreo",
                        "Usuario Correo",
                        "Correo",
                        "Email",
                        "CorreoUsuario"
                    ],
                    mapaColumnas
                ),

                obtenerCampoFlexible(
                    item,
                    [
                        "Title"
                    ],
                    mapaColumnas
                )
            ];

            const correosNormalizados =
                posiblesCorreos
                    .map(
                        valor =>
                            extraerCorreo(
                                valor
                            )
                    )
                    .filter(
                        Boolean
                    );

            console.log(
                "Comparando:",
                correo,
                "CONTRA:",
                correosNormalizados
            );

            if (
                correosNormalizados.includes(
                    correo
                )
            ) {

                registro =
                    item;

                break;
            }
        }

        /* =====================================================
           SIN REGISTRO
        ===================================================== */

        if (
            !registro
        ) {

            console.warn(
                "========================================="
            );

            console.warn(
                "NO SE ENCONTRÓ REGISTRO PARA:",
                correo
            );

            console.warn(
                "========================================="
            );

            return crearPermisosDenegados(
                correo,
                cuenta.name || correo
            );
        }

        /* =====================================================
           REGISTRO ENCONTRADO
        ===================================================== */

        console.log(
            "========================================="
        );

        console.log(
            "REGISTRO DEL USUARIO ENCONTRADO:"
        );

        console.log(
            registro
        );

        console.log(
            "CAMPOS DEL REGISTRO:"
        );

        console.log(
            registro.fields
        );

        console.log(
            "========================================="
        );

        /* =====================================================
           GRUPO
        ===================================================== */

        const grupo =
            String(
                obtenerCampoFlexible(
                    registro,
                    [
                        "Grupo"
                    ],
                    mapaColumnas
                ) ||
                "Normal"
            ).trim();

        console.log(
            "GRUPO DEL USUARIO:",
            grupo
        );

        /* =====================================================
           GRUPO TI
        ===================================================== */

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
                cuenta.name || correo
            );
        }

        /* =====================================================
           PERMISOS INDIVIDUALES
        ===================================================== */

        const permisos = {

            Oficina:
                valorBooleano(
                    obtenerCampoFlexible(
                        registro,
                        [
                            "Oficina"
                        ],
                        mapaColumnas
                    )
                ),

            Personal:
                valorBooleano(
                    obtenerCampoFlexible(
                        registro,
                        [
                            "Personal"
                        ],
                        mapaColumnas
                    )
                ),

            Reservas:
                valorBooleano(
                    obtenerCampoFlexible(
                        registro,
                        [
                            "Reservas"
                        ],
                        mapaColumnas
                    )
                ),

            Salas:
                valorBooleano(
                    obtenerCampoFlexible(
                        registro,
                        [
                            "Salas"
                        ],
                        mapaColumnas
                    )
                ),

            Comunicados:
                valorBooleano(
                    obtenerCampoFlexible(
                        registro,
                        [
                            "Comunicados"
                        ],
                        mapaColumnas
                    )
                ),

            Seguridad:
                valorBooleano(
                    obtenerCampoFlexible(
                        registro,
                        [
                            "Seguridad"
                        ],
                        mapaColumnas
                    )
                ),

            Infraestructura:
                valorBooleano(
                    obtenerCampoFlexible(
                        registro,
                        [
                            "Infraestructura"
                        ],
                        mapaColumnas
                    )
                ),

            Tickets:
                valorBooleano(
                    obtenerCampoFlexible(
                        registro,
                        [
                            "Tickets"
                        ],
                        mapaColumnas
                    )
                ),

            Configuracion:
                valorBooleano(
                    obtenerCampoFlexible(
                        registro,
                        [
                            "Configuracion",
                            "Configuración"
                        ],
                        mapaColumnas
                    )
                ),

            /*
             * Un usuario Normal nunca puede administrar
             * permisos.
             */

            Permisos:
                false
        };

        console.log(
            "========================================="
        );

        console.log(
            "PERMISOS LEÍDOS DESDE SHAREPOINT:"
        );

        console.table(
            permisos
        );

        console.log(
            "========================================="
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
                    ],
                    mapaColumnas
                ) ||
                correo,

            grupo:
                "Normal",

            permisos:
                permisos
        };
    }

    /* =========================================================
       CONSTRUIR MAPA DE COLUMNAS SHAREPOINT
    ========================================================= */

    function construirMapaColumnas (
        columnas
    ) {

        const mapa = {};

        columnas.forEach(
            columna => {

                const nombreInterno =
                    columna.name || "";

                const nombreVisible =
                    columna.displayName || "";

                if (
                    nombreInterno
                ) {

                    mapa[
                        normalizarTexto(
                            nombreInterno
                        )
                    ] =
                        nombreInterno;
                }

                if (
                    nombreVisible
                ) {

                    mapa[
                        normalizarTexto(
                            nombreVisible
                        )
                    ] =
                        nombreInterno;
                }
            }
        );

        return mapa;
    }

    /* =========================================================
       EXTRAER CORREO
    ========================================================= */

    function extraerCorreo (
        valor
    ) {

        if (
            !valor
        ) {

            return "";
        }

        if (
            typeof valor === "string"
        ) {

            const texto =
                valor.trim();

            const coincidencia =
                texto.match(
                    /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i
                );

            return normalizarCorreo(
                coincidencia
                    ? coincidencia[0]
                    : texto
            );
        }

        if (
            typeof valor === "object"
        ) {

            const posibles = [

                valor.email,

                valor.Email,

                valor.mail,

                valor.Mail,

                valor.userPrincipalName,

                valor.UserPrincipalName,

                valor.username,

                valor.Username,

                valor.text,

                valor.displayName
            ];

            for (
                const posible of posibles
            ) {

                const correo =
                    extraerCorreo(
                        posible
                    );

                if (
                    correo &&
                    correo.includes("@")
                ) {

                    return correo;
                }
            }
        }

        return "";
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
            "APLICANDO PERMISOS AL MENÚ:",
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

        ruta =
            ruta.replace(
                /^\/live-office\/?/i,
                ""
            );

        ruta =
            ruta.replace(
                /^\/+/,
                ""
            );

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

        if (
            paginaActual ===
            "index.html"
        ) {

            mostrarBloqueoOficina(
                modulo
            );

            return;
        }

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
        nombre,
        mapaColumnas = {}
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

        /*
         * Primero intentamos con el nombre interno
         * resuelto mediante /columns.
         */

        const nombreInterno =
            mapaColumnas[
                objetivo
            ];

        if (
            nombreInterno &&
            Object.prototype.hasOwnProperty.call(
                fields,
                nombreInterno
            )
        ) {

            return fields[
                nombreInterno
            ];
        }

        /*
         * Después buscamos directamente entre
         * las propiedades recibidas por Graph.
         */

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
        nombres,
        mapaColumnas = {}
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
                    nombre,
                    mapaColumnas
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
            error.codigo ===
            "MSAL_NO_DISPONIBLE"
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

                    cuenta.idTokenClaims?.upn ||

                    ""
                );
            },

        limpiarSesion:
            limpiarSesion
    };

})();