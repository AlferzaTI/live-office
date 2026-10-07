/* =========================================================
   PROTECCIÓN DE ACCESO
   ALFERZA LIVE OFFICE
========================================================= */

(function protegerPagina() {

    /* =========================================================
       CONFIGURACIÓN
    ========================================================= */

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


    /* =========================================================
       SCOPES
    ========================================================= */

    /*
       IMPORTANTE:

       login.js utiliza Sites.Read.All.
       auth.js NO debe solicitar Sites.ReadWrite.All.
    */

    const SCOPES = [
        "User.Read",
        "Sites.Read.All"
    ];


    /* =========================================================
       PÁGINAS -> PERMISOS
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
       MÓDULOS
    ========================================================= */

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


    /* =========================================================
       COMPROBAR SESIÓN
    ========================================================= */

    const rutaActual =
        window.location.pathname;

    /*
       login.html y blank.html no necesitan protección.
    */

    if (
        rutaActual === LOGIN_PATH ||
        rutaActual === BLANK_PATH
    ) {
        return;
    }


    const sesion =
        sessionStorage.getItem(
            SESION_KEY
        );


    /*
       Si no existe la sesión creada por login.js,
       regresar al login.
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
       INICIAR
    ========================================================= */

    iniciarControlPermisos();


    /* =========================================================
       GUARDAR URL DE RETORNO
    ========================================================= */

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


    /* =========================================================
       INICIAR CONTROL DE PERMISOS
    ========================================================= */

    async function iniciarControlPermisos() {

        try {

            /*
               Esperar a que exista el DOM.
            */

            if (
                document.readyState === "loading"
            ) {

                await new Promise(resolve => {

                    document.addEventListener(
                        "DOMContentLoaded",
                        resolve,
                        {
                            once: true
                        }
                    );

                });

            }


            console.log(
                "========================================"
            );

            console.log(
                "ALFERZA LIVE OFFICE"
            );

            console.log(
                "Verificando permisos..."
            );

            console.log(
                "========================================"
            );


            /*
               Obtener permisos directamente desde
               SharePoint.
            */

            const datos =
                await obtenerPermisosUsuario();


            console.log(
                "DATOS FINALES DE PERMISOS:",
                datos
            );


            /*
               Guardar resultado actual.
            */

            sessionStorage.setItem(
                PERMISOS_KEY,
                JSON.stringify(datos)
            );


            /*
               Aplicar menú.
            */

            aplicarPermisosMenu(
                datos
            );


            /*
               Validar página actual.
            */

            validarPaginaActual(
                datos
            );


        }
        catch (error) {

            console.error(
                "========================================"
            );

            console.error(
                "ERROR COMPROBANDO PERMISOS"
            );

            console.error(
                error
            );

            console.error(
                "========================================"
            );


            /*
               Por seguridad:

               Si no podemos comprobar los permisos,
               NO otorgamos acceso.
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


    /* =========================================================
       MSAL
    ========================================================= */

    let instanciaMSAL = null;


    function crearMSAL() {

        if (
            typeof msal === "undefined"
        ) {

            throw new Error(
                "MSAL no está disponible."
            );
        }


        if (
            instanciaMSAL
        ) {

            return instanciaMSAL;
        }


        instanciaMSAL =
            new msal.PublicClientApplication({

                auth: {

                    clientId:
                        CLIENT_ID,

                    authority:
                        AUTHORITY,

                    redirectUri:
                        REDIRECT_URI
                },

                cache: {

                    cacheLocation:
                        "sessionStorage",

                    storeAuthStateInCookie:
                        false
                }

            });


        return instanciaMSAL;
    }


    /* =========================================================
       CUENTA ACTUAL
    ========================================================= */

    async function obtenerCuentaActual() {

        const instancia =
            crearMSAL();


        const cuentas =
            instancia.getAllAccounts();


        console.log(
            "Cuentas MSAL encontradas:",
            cuentas
        );


        if (
            !cuentas ||
            cuentas.length === 0
        ) {

            throw new Error(
                "No existe una cuenta de Microsoft en la sesión."
            );
        }


        /*
           Buscar específicamente una cuenta
           @alferza.pe.
        */

        const cuentaAlferza =
            cuentas.find(
                cuenta => {

                    const correo =
                        normalizarCorreo(

                            cuenta.username ||

                            cuenta.idTokenClaims?.preferred_username ||

                            cuenta.idTokenClaims?.email ||

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


        console.log(
            "Cuenta seleccionada:",
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
       TOKEN
    ========================================================= */

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


            if (
                !resultado ||
                !resultado.accessToken
            ) {

                throw new Error(
                    "Microsoft no devolvió un access token."
                );
            }


            console.log(
                "Token Microsoft obtenido correctamente."
            );


            return resultado.accessToken;


        }
        catch (error) {

            console.error(
                "Error obteniendo token silencioso:",
                error
            );


            throw new Error(
                "No se pudo obtener el token de Microsoft para consultar PermisosTI."
            );
        }
    }


    /* =========================================================
       GRAPH FETCH
    ========================================================= */

    async function graphFetch(
        url,
        token
    ) {

        console.log(
            "Graph:",
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


            console.error(
                "Graph error:",
                respuesta.status,
                texto
            );


            throw new Error(
                `Graph ${respuesta.status}: ${texto}`
            );
        }


        return respuesta.json();
    }


    /* =========================================================
       OBTENER PERMISOS DEL USUARIO
    ========================================================= */

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
           Obtener correo.
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


        console.log(
            "========================================"
        );

        console.log(
            "USUARIO ACTUAL:",
            correo
        );

        console.log(
            "========================================"
        );


        /* =====================================================
           TI INICIAL
        ===================================================== */

        if (
            correo === USUARIO_TI_INICIAL
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


        /* =====================================================
           OBTENER SITIO
        ===================================================== */

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
            "Sitio SharePoint encontrado:",
            sitio
        );


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
            "Listas encontradas:",
            listas.value
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
                        displayName === objetivo ||
                        name === objetivo
                    );

                }

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
                resultado[
                    "@odata.nextLink"
                ] ||
                null;
        }


        console.log(
            "Cantidad de registros en PermisosTI:",
            items.length
        );


        /*
           Mostrar campos de cada registro en consola.

           Esto es importante para detectar si
           SharePoint está devolviendo nombres internos
           diferentes.
        */

        console.log(
            "Registros obtenidos de PermisosTI:",
            items
        );


        /* =====================================================
           BUSCAR USUARIO
        ===================================================== */

        const registro =
            items.find(

                item => {

                    const correoRegistro =
                        obtenerCampoFlexible(
                            item,
                            [
                                "UsuarioCorreo",
                                "UsuarioCorreo0",
                                "Correo",
                                "Email",
                                "Usuario"
                            ]
                        );


                    const correoNormalizado =
                        normalizarCorreo(
                            extraerCorreo(
                                correoRegistro
                            )
                        );


                    console.log(
                        "Comparando usuario:",
                        correoNormalizado,
                        "vs",
                        correo
                    );


                    return (
                        correoNormalizado ===
                        correo
                    );

                }

            );


        /* =====================================================
           USUARIO NO ENCONTRADO
        ===================================================== */

        if (!registro) {

            console.warn(
                "========================================"
            );

            console.warn(
                "USUARIO NO ENCONTRADO EN PermisosTI"
            );

            console.warn(
                "Correo buscado:",
                correo
            );

            console.warn(
                "========================================"
            );


            return crearPermisosDenegados(
                correo,
                cuenta.name ||
                correo
            );
        }


        console.log(
            "========================================"
        );

        console.log(
            "REGISTRO DEL USUARIO ENCONTRADO"
        );

        console.log(
            registro
        );

        console.log(
            "FIELDS:",
            registro.fields
        );

        console.log(
            "========================================"
        );


        /* =====================================================
           GRUPO
        ===================================================== */

        const grupoValor =
            obtenerCampoFlexible(
                registro,
                [
                    "Grupo"
                ]
            );


        const grupo =
            String(
                grupoValor ||
                "Normal"
            ).trim();


        console.log(
            "Grupo del usuario:",
            grupo
        );


        /* =====================================================
           GRUPO TI
        ===================================================== */

        if (
            normalizarTexto(grupo) ===
            normalizarTexto("TI")
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


        /* =====================================================
           USUARIO NORMAL
        ===================================================== */

        const permisos = {};


        MODULOS.forEach(
            modulo => {

                permisos[modulo] =
                    false;

            }
        );


        /*
           Leer cada columna de permisos
           directamente desde SharePoint.
        */

        MODULOS.forEach(
            modulo => {

                const valor =
                    obtenerCampoFlexible(
                        registro,
                        obtenerAliasesModulo(
                            modulo
                        )
                    );


                permisos[modulo] =
                    valorBooleano(
                        valor
                    );


                console.log(
                    `Permiso ${modulo}:`,
                    valor,
                    "=>",
                    permisos[modulo]
                );

            }
        );


        console.log(
            "========================================"
        );

        console.log(
            "PERMISOS FINALES DEL USUARIO:"
        );

        console.table(
            permisos
        );

        console.log(
            "========================================"
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
                        "Nombre",
                        "Title"
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
       ALIASES DE MÓDULOS
    ========================================================= */

    function obtenerAliasesModulo(
        modulo
    ) {

        const aliases = {

            "Oficina": [
                "Oficina",
                "oficina"
            ],

            "Personal": [
                "Personal",
                "personal"
            ],

            "Reservas": [
                "Reservas",
                "reservas",
                "Reserva"
            ],

            "Salas": [
                "Salas",
                "salas",
                "Sala"
            ],

            "Comunicados": [
                "Comunicados",
                "comunicados"
            ],

            "Seguridad": [
                "Seguridad",
                "seguridad"
            ],

            "Infraestructura": [
                "Infraestructura",
                "infraestructura"
            ],

            "Tickets": [
                "Tickets",
                "tickets"
            ],

            "Permisos": [
                "Permisos",
                "permisos"
            ],

            "Configuracion": [
                "Configuracion",
                "Configuración",
                "configuracion"
            ]
        };


        return (
            aliases[modulo] ||
            [modulo]
        );
    }


    /* =========================================================
       CREAR PERMISOS TI
    ========================================================= */

    function crearPermisosTI(
        correo,
        nombre
    ) {

        const permisos = {};


        MODULOS.forEach(
            modulo => {

                permisos[modulo] =
                    true;

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


    /* =========================================================
       CREAR PERMISOS DENEGADOS
    ========================================================= */

    function crearPermisosDenegados(
        correo = "",
        nombre = ""
    ) {

        const permisos = {};


        MODULOS.forEach(
            modulo => {

                permisos[modulo] =
                    false;

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


    /* =========================================================
       APLICAR PERMISOS AL MENÚ
    ========================================================= */

    function aplicarPermisosMenu(
        datos
    ) {

        const permisos =
            datos?.permisos ||
            {};


        const elementos =
            document.querySelectorAll(
                ".sidebar li"
            );


        console.log(
            "Elementos del menú encontrados:",
            elementos.length
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
                    permisos[modulo] === true;


                elemento.style.display =
                    permitido
                        ? ""
                        : "none";


                console.log(
                    `Menú ${modulo}:`,
                    permitido
                );

            }
        );
    }


    /* =========================================================
       OBTENER NOMBRE DEL MÓDULO
    ========================================================= */

    function obtenerNombreModulo(
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

    function validarPaginaActual(
        datos
    ) {

        const pagina =
            obtenerPaginaActual();


        console.log(
            "Página actual:",
            pagina
        );


        /*
           Raíz sin página específica.
        */

        if (!pagina) {
            return;
        }


        if (
            pagina === "login.html"
        ) {
            return;
        }


        const modulo =
            encontrarPermisoPagina(
                pagina
            );


        /*
           Página no registrada.
        */

        if (!modulo) {

            console.log(
                "Página sin permiso específico:",
                pagina
            );

            return;
        }


        const permitido =
            datos?.permisos?.[modulo] === true;


        console.log(
            `Validación ${modulo}:`,
            permitido
        );


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


    /* =========================================================
       OBTENER PÁGINA ACTUAL
    ========================================================= */

    function obtenerPaginaActual() {

        let ruta =
            window.location.pathname;


        const prefijo =
            ROOT_PATH;


        if (
            ruta.startsWith(prefijo)
        ) {

            ruta =
                ruta.substring(
                    prefijo.length
                );
        }


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


    /* =========================================================
       BUSCAR PERMISO DE PÁGINA
    ========================================================= */

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


    /* =========================================================
       BLOQUEAR PÁGINA
    ========================================================= */

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


    /* =========================================================
       OBTENER CAMPO SHAREPOINT
    ========================================================= */

    function obtenerCampo(
        registro,
        nombre
    ) {

        return obtenerCampoFlexible(
            registro,
            [nombre]
        );
    }


    /* =========================================================
       OBTENER CAMPO FLEXIBLE
    ========================================================= */

    function obtenerCampoFlexible(
        registro,
        nombres
    ) {

        if (
            !registro ||
            !registro.fields
        ) {

            return "";
        }


        const fields =
            registro.fields;


        const listaNombres =
            Array.isArray(nombres)
                ? nombres
                : [nombres];


        /*
           Primero buscar coincidencia exacta.
        */

        for (
            const nombre
            of listaNombres
        ) {

            if (
                Object.prototype.hasOwnProperty.call(
                    fields,
                    nombre
                )
            ) {

                return fields[nombre];
            }
        }


        /*
           Después buscar ignorando:

           - mayúsculas
           - minúsculas
           - espacios
           - guiones
           - acentos
        */

        const claves =
            Object.keys(
                fields
            );


        for (
            const nombre
            of listaNombres
        ) {

            const objetivo =
                normalizarTexto(
                    nombre
                );


            const clave =
                claves.find(
                    key =>
                        normalizarTexto(
                            key
                        ) === objetivo
                );


            if (
                clave
            ) {

                return fields[
                    clave
                ];
            }
        }


        return "";
    }


    /* =========================================================
       EXTRAER CORREO
    ========================================================= */

    function extraerCorreo(
        valor
    ) {

        if (
            !valor
        ) {

            return "";
        }


        /*
           Si SharePoint devuelve un objeto.
        */

        if (
            typeof valor === "object"
        ) {

            return (

                valor.email ||

                valor.mail ||

                valor.userPrincipalName ||

                valor.username ||

                ""

            );
        }


        return String(
            valor
        );
    }


    /* =========================================================
       VALOR BOOLEANO
    ========================================================= */

    function valorBooleano(
        valor
    ) {

        if (
            valor === true
        ) {

            return true;
        }


        if (
            valor === false
        ) {

            return false;
        }


        if (
            valor === 1
        ) {

            return true;
        }


        if (
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


        if (
            [
                "true",
                "1",
                "yes",
                "si",
                "sí",
                "on",
                "checked",
                "activo",
                "permitido"
            ].includes(texto)
        ) {

            return true;
        }


        return false;
    }


    /* =========================================================
       NORMALIZAR TEXTO
    ========================================================= */

    function normalizarTexto(
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

    function normalizarCorreo(
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

    function normalizarRuta(
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