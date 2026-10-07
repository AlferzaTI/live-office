(function () {

    /* =========================================================
       ALFERZA LIVE OFFICE
       GESTIÓN DE PERMISOS
    ========================================================= */


    /* =========================================================
       CONFIGURACIÓN MICROSOFT
       AUTH.JS ES EL RESPONSABLE DE MSAL
    ========================================================= */

    const SHAREPOINT_HOST =
        "alferzaholding-my.sharepoint.com";


    const SHAREPOINT_SITE_PATH =
        "/personal/soporte1_alferza_pe";


    const PERMISOS_LIST_NAME =
        "PermisosTI";


    /* =========================================================
       SCOPES GRAPH
    ========================================================= */

    const SCOPES_GRAPH = [

        "User.Read",

        "Sites.ReadWrite.All"

    ];


    /* =========================================================
       GRUPOS
    ========================================================= */

    const GRUPO_TI =
        "TI";


    const GRUPO_NORMAL =
        "Normal";


    /* =========================================================
       USUARIOS TI INICIALES
       SOLO CUENTAS DE EMERGENCIA / ADMINISTRACIÓN INICIAL
    ========================================================= */

    const USUARIOS_TI_INICIALES = [

        "soporte1@alferza.pe"

    ];


    /* =========================================================
       MÓDULOS
    ========================================================= */

    const MODULOS_PERMISOS = [

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


    /* =========================================================
       VALIDAR AUTH.JS
    ========================================================= */

    if (
        !window.alferzaAuth ||
        !window.alferzaMsal
    ) {

        mostrarErrorInicial(
            "No se pudo inicializar la autenticación de Microsoft."
        );


        console.error(
            "ALFERZA LIVE OFFICE: auth.js debe cargarse antes que permisos.js."
        );


        return;

    }


    /* =========================================================
       MSAL
       SE UTILIZA LA INSTANCIA CREADA POR AUTH.JS
    ========================================================= */

    const msalInstance =
        window.alferzaMsal;


    /* =========================================================
       VARIABLES
    ========================================================= */

    let cuentaActual =
        null;


    let sitioSharePoint =
        null;


    let listaPermisos =
        null;


    let usuariosGraph =
        [];


    let registrosPermisos =
        [];


    let usuarioSeleccionado =
        null;


    let mapaCampos =
        null;


    /* =========================================================
       ELEMENTOS
    ========================================================= */

    const selectorUsuario =
        document.getElementById(
            "selectorUsuario"
        );


    const correoUsuario =
        document.getElementById(
            "correoUsuario"
        );


    const btnBuscarUsuario =
        document.getElementById(
            "btnBuscarUsuario"
        );


    const panelPermisos =
        document.getElementById(
            "panelPermisos"
        );


    const nombreUsuario =
        document.getElementById(
            "nombreUsuario"
        );


    const correoMostrado =
        document.getElementById(
            "correoMostrado"
        );


    const grupoUsuario =
        document.getElementById(
            "grupoUsuario"
        );


    const btnGuardarPermisos =
        document.getElementById(
            "btnGuardarPermisos"
        );


    const tablaPermisos =
        document.getElementById(
            "tablaPermisos"
        );


    const userInfo =
        document.getElementById(
            "userInfo"
        );


    /* =========================================================
       INICIO
    ========================================================= */

    document.addEventListener(
        "DOMContentLoaded",
        iniciarPermisos
    );


    /* =========================================================
       MOSTRAR USUARIO EN NAVBAR
    ========================================================= */

    function mostrarUsuarioNavbar() {

        if (
            !userInfo ||
            !cuentaActual
        ) {

            return;

        }


        userInfo.textContent =

            cuentaActual.name ||

            cuentaActual.username ||

            "";

    }


    /* =========================================================
       INICIAR
    ========================================================= */

    async function iniciarPermisos() {

        console.log(
            "Iniciando módulo Permisos..."
        );


        configurarEventos();


        ocultarPanel();


        try {

            /* =================================================
               ESPERAR AUTENTICACIÓN DE AUTH.JS
            ================================================= */

            if (
                window.alferzaAuth &&
                typeof window.alferzaAuth.inicializar ===
                    "function"
            ) {

                await window.alferzaAuth.inicializar();

            }


            /* =================================================
               OBTENER CUENTA
            ================================================= */

            const cuenta =
                obtenerCuenta();


            if (!cuenta) {

                mostrarMensaje(

                    "No existe una sesión de Microsoft.",

                    "error"

                );


                return;

            }


            cuentaActual =
                cuenta;


            mostrarUsuarioNavbar();


            console.log(
                "Cuenta autenticada:",
                cuentaActual
            );


            /* =================================================
               CORREO ACTUAL
            ================================================= */

            const correoActual =
                obtenerCorreoCuentaActual();


            console.log(
                "Correo actual:",
                correoActual
            );


            /* =================================================
               TI INICIAL
            ================================================= */

            const esTIInicial =

                USUARIOS_TI_INICIALES
                    .map(
                        normalizarCorreo
                    )
                    .includes(
                        normalizarCorreo(
                            correoActual
                        )
                    );


            console.log(
                "¿Es TI inicial?:",
                esTIInicial
            );


            /* =================================================
               CARGAR USUARIOS
            ================================================= */

            await cargarUsuariosGraph();


            /* =================================================
               CARGAR PERMISOS
            ================================================= */

            try {

                await cargarPermisos();

            }

            catch (errorLista) {

                console.error(

                    "No se pudo cargar PermisosTI:",

                    errorLista

                );


                /*
                 * El administrador inicial puede continuar
                 * aunque la lista tenga un problema.
                 *
                 * Para cualquier otro usuario NO concedemos
                 * permisos si no podemos consultar PermisosTI.
                 */

                if (esTIInicial) {

                    console.warn(

                        "El usuario TI inicial continúa " +
                        "aunque PermisosTI no esté disponible."

                    );


                    return;

                }


                mostrarAccesoDenegado();


                return;

            }


            /* =================================================
               SI ES TI INICIAL
            ================================================= */

            if (esTIInicial) {

                console.log(
                    "Usuario autorizado como TI inicial."
                );


                return;

            }


            /* =================================================
               BUSCAR REGISTRO DEL USUARIO ACTUAL
            ================================================= */

            const registroActual =
                buscarRegistroPorCorreo(
                    correoActual
                );


            console.log(
                "Registro PermisosTI del usuario actual:",
                registroActual
            );


            /* =================================================
               COMPROBAR GRUPO
            ================================================= */

            const grupoActual =

                registroActual

                    ? normalizarTexto(

                        obtenerCampo(
                            registroActual,
                            "Grupo"
                        )

                    )

                    : "";


            const esTILista =

                grupoActual ===
                normalizarTexto(
                    GRUPO_TI
                );


            console.log(
                "Grupo del usuario actual:",
                grupoActual
            );


            console.log(
                "¿Autorizado por PermisosTI?:",
                esTILista
            );


            /* =================================================
               DENEGAR
            ================================================= */

            if (!esTILista) {

                mostrarAccesoDenegado();


                return;

            }


            /* =================================================
               AUTORIZADO
            ================================================= */

            console.log(
                "Usuario autorizado para administrar Permisos."
            );

        }

        catch (error) {

            console.error(

                "Error inicializando Permisos:",

                error

            );


            mostrarMensaje(

                obtenerMensajeError(
                    error
                ),

                "error"

            );

        }

    }


    /* =========================================================
       CONFIGURAR EVENTOS
    ========================================================= */

    function configurarEventos() {


        if (
            btnBuscarUsuario
        ) {

            btnBuscarUsuario.addEventListener(

                "click",

                buscarUsuario

            );

        }


        if (
            correoUsuario
        ) {

            correoUsuario.addEventListener(

                "keydown",

                function (event) {

                    if (
                        event.key ===
                        "Enter"
                    ) {

                        event.preventDefault();

                        buscarUsuario();

                    }

                }

            );

        }


        if (
            selectorUsuario
        ) {

            selectorUsuario.addEventListener(

                "change",

                function () {

                    seleccionarUsuario(
                        this.value
                    );

                }

            );

        }


        if (
            grupoUsuario
        ) {

            grupoUsuario.addEventListener(

                "change",

                cambiarGrupo

            );

        }


        if (
            btnGuardarPermisos
        ) {

            btnGuardarPermisos.addEventListener(

                "click",

                guardarPermisos

            );

        }


        configurarMenu();

    }


    /* =========================================================
       MENÚ
    ========================================================= */

    function configurarMenu() {

        const menuToggle =
            document.getElementById(
                "menuToggle"
            );


        const sidebar =
            document.getElementById(
                "sidebar"
            );


        if (
            menuToggle &&
            sidebar
        ) {

            menuToggle.addEventListener(

                "click",

                function () {

                    sidebar.classList.toggle(
                        "open"
                    );

                }

            );

        }

    }


    /* =========================================================
       OBTENER CUENTA
    ========================================================= */

    function obtenerCuenta() {

        /*
         * Primero utilizamos auth.js.
         */

        if (
            window.alferzaAuth &&
            typeof window.alferzaAuth.obtenerCuenta ===
                "function"
        ) {

            const cuentaAuth =
                window.alferzaAuth.obtenerCuenta();


            if (cuentaAuth) {

                cuentaActual =
                    cuentaAuth;


                return cuentaAuth;

            }

        }


        /*
         * Respaldo usando la instancia compartida.
         */

        const cuentas =
            msalInstance.getAllAccounts();


        if (
            cuentas &&
            cuentas.length > 0
        ) {

            cuentaActual =
                msalInstance.getActiveAccount() ||
                cuentas[0];


            msalInstance.setActiveAccount(
                cuentaActual
            );


            return cuentaActual;

        }


        return null;

    }


    /* =========================================================
       TOKEN
       UTILIZA AUTH.JS
    ========================================================= */

    async function obtenerToken() {

        /*
         * Utilizamos el mecanismo centralizado
         * de auth.js.
         */

        if (
            window.alferzaAuth &&
            typeof window.alferzaAuth.obtenerToken ===
                "function"
        ) {

            return await window.alferzaAuth.obtenerToken(
                SCOPES_GRAPH
            );

        }


        /*
         * Respaldo.
         */

        const cuenta =
            obtenerCuenta();


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
                        SCOPES_GRAPH,

                    account:
                        cuenta

                });


            return resultado.accessToken;

        }

        catch (error) {

            console.error(

                "No se pudo obtener el token silenciosamente:",

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


    /* =========================================================
       GRAPH FETCH
    ========================================================= */

    async function graphFetch(
        url,
        opciones = {}
    ) {

        const token =
            await obtenerToken();


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


    /* =========================================================
       OBTENER USUARIOS GRAPH
    ========================================================= */

    async function cargarUsuariosGraph() {

        if (
            selectorUsuario
        ) {

            selectorUsuario.innerHTML =

                `
                <option value="">
                    Cargando usuarios de Microsoft...
                </option>
                `;

        }


        let url =

            "https://graph.microsoft.com/v1.0/users" +

            "?$select=id,displayName,mail,userPrincipalName,accountEnabled" +

            "&$top=999";


        const usuarios =
            [];


        while (url) {

            const resultado =
                await graphFetch(
                    url
                );


            if (
                resultado &&
                Array.isArray(
                    resultado.value
                )
            ) {

                usuarios.push(
                    ...resultado.value
                );

            }


            url =
                resultado[
                    "@odata.nextLink"
                ] ||
                null;

        }


        /*
         * Solo usuarios ALFERZA.
         */

        usuariosGraph =
            usuarios
                .filter(

                    usuario => {

                        const correo =

                            usuario.mail ||

                            usuario.userPrincipalName ||

                            "";


                        return normalizarCorreo(
                            correo
                        ).endsWith(
                            "@alferza.pe"
                        );

                    }

                )
                .sort(

                    (a, b) => {

                        const nombreA =
                            String(
                                a.displayName ||
                                ""
                            )
                                .toLowerCase();


                        const nombreB =
                            String(
                                b.displayName ||
                                ""
                            )
                                .toLowerCase();


                        return nombreA.localeCompare(

                            nombreB,

                            "es"

                        );

                    }

                );


        renderizarSelectorUsuarios(
            usuariosGraph
        );


        console.log(

            "Usuarios ALFERZA obtenidos:",

            usuariosGraph.length

        );

    }


    /* =========================================================
       RENDERIZAR SELECTOR
    ========================================================= */

    function renderizarSelectorUsuarios(
        usuarios
    ) {

        if (
            !selectorUsuario
        ) {

            return;

        }


        selectorUsuario.innerHTML =
            "";


        const opcionInicial =
            document.createElement(
                "option"
            );


        opcionInicial.value =
            "";


        opcionInicial.textContent =

            `Selecciona un usuario (${usuarios.length})`;


        selectorUsuario.appendChild(
            opcionInicial
        );


        usuarios.forEach(

            usuario => {

                const correo =

                    usuario.mail ||

                    usuario.userPrincipalName ||

                    "";


                const opcion =
                    document.createElement(
                        "option"
                    );


                opcion.value =
                    usuario.id;


                opcion.textContent =

                    `${usuario.displayName || "Sin nombre"} — ${correo}`;


                selectorUsuario.appendChild(
                    opcion
                );

            }

        );


        if (
            !usuarios.length
        ) {

            opcionInicial.textContent =
                "No se encontraron usuarios.";

        }

    }


    /* =========================================================
       BUSCAR USUARIO
    ========================================================= */

    async function buscarUsuario() {

        const texto =
            normalizarTexto(

                correoUsuario
                    ? correoUsuario.value
                    : ""

            );


        if (!texto) {

            renderizarSelectorUsuarios(
                usuariosGraph
            );


            return;

        }


        const resultados =
            usuariosGraph.filter(

                usuario => {

                    const nombre =

                        normalizarTexto(
                            usuario.displayName
                        );


                    const correo =

                        normalizarTexto(

                            usuario.mail ||

                            usuario.userPrincipalName ||

                            ""

                        );


                    return (

                        nombre.includes(
                            texto
                        ) ||

                        correo.includes(
                            texto
                        )

                    );

                }

            );


        renderizarSelectorUsuarios(
            resultados
        );


        if (
            resultados.length ===
            1
        ) {

            selectorUsuario.value =
                resultados[0].id;


            seleccionarUsuario(
                resultados[0].id
            );

        }

        else if (
            resultados.length ===
            0
        ) {

            mostrarMensaje(

                "No se encontró ningún usuario.",

                "error"

            );

        }

        else {

            mostrarMensaje(

                `${resultados.length} usuarios encontrados. Selecciona uno de la lista.`,

                "info"

            );

        }

    }


    /* =========================================================
       SELECCIONAR USUARIO
    ========================================================= */

    function seleccionarUsuario(
        usuarioId
    ) {

        if (!usuarioId) {

            usuarioSeleccionado =
                null;


            ocultarPanel();


            return;

        }


        const usuario =
            usuariosGraph.find(

                item =>

                    String(
                        item.id
                    ) ===
                    String(
                        usuarioId
                    )

            );


        if (!usuario) {

            return;

        }


        const correo =

            normalizarCorreo(

                usuario.mail ||

                usuario.userPrincipalName ||

                ""

            );


        usuarioSeleccionado = {

            id:
                usuario.id,

            nombre:
                usuario.displayName ||
                "Sin nombre",

            correo:
                correo,

            userPrincipalName:
                usuario.userPrincipalName ||
                ""

        };


        console.log(

            "Usuario seleccionado:",

            usuarioSeleccionado

        );


        if (nombreUsuario) {

            nombreUsuario.textContent =

                usuarioSeleccionado.nombre;

        }


        if (correoMostrado) {

            correoMostrado.textContent =

                usuarioSeleccionado.correo;

        }


        if (correoUsuario) {

            correoUsuario.value =

                usuarioSeleccionado.correo;

        }


        /*
         * Buscar configuración guardada.
         */

        const registro =
            buscarRegistroPorCorreo(

                usuarioSeleccionado.correo

            );


        if (registro) {

            cargarRegistroEnFormulario(
                registro
            );

        }

        else {

            prepararNuevoUsuario();

        }


        mostrarPanel();

    }


    /* =========================================================
       MOSTRAR PANEL
    ========================================================= */

    function mostrarPanel() {

        if (
            panelPermisos
        ) {

            panelPermisos.style.display =
                "block";

        }

    }


    /* =========================================================
       OCULTAR PANEL
    ========================================================= */

    function ocultarPanel() {

        if (
            panelPermisos
        ) {

            panelPermisos.style.display =
                "none";

        }

    }


    /* =========================================================
       OBTENER SITIO SHAREPOINT
    ========================================================= */

    async function obtenerSitioSharePoint() {

        if (
            sitioSharePoint
        ) {

            return sitioSharePoint;

        }


        const url =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;


        console.log(
            "Obteniendo sitio SharePoint:",
            url
        );


        sitioSharePoint =
            await graphFetch(
                url
            );


        console.log(
            "Sitio SharePoint encontrado:",
            sitioSharePoint
        );


        return sitioSharePoint;

    }


    /* =========================================================
       OBTENER LISTA PERMISOS
    ========================================================= */

    async function obtenerListaPermisos() {

    if (
        listaPermisos
    ) {

        return listaPermisos;

    }


    const sitio =
        await obtenerSitioSharePoint();


    if (
        !sitio ||
        !sitio.id
    ) {

        throw new Error(
            "No se pudo obtener el ID del sitio SharePoint."
        );

    }


    console.log(
        "Buscando lista:",
        PERMISOS_LIST_NAME
    );


    /* =====================================================
       MÉTODO 1
       ENUMERAR LISTAS
    ===================================================== */

    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists` +
        `?$select=id,name,displayName&$top=200`;


    const resultado =
        await graphFetch(
            url
        );


    const listas =

        resultado &&
        Array.isArray(
            resultado.value
        )

            ? resultado.value

            : [];


    console.log(
        "Listas devueltas por Microsoft Graph:",
        listas
    );


    /* =====================================================
       BUSCAR PERMISOSTI
    ===================================================== */

    const objetivo =
        normalizarTexto(
            PERMISOS_LIST_NAME
        );


    listaPermisos =
        listas.find(

            lista => {

                const displayName =
                    normalizarTexto(
                        lista.displayName
                    );


                const name =
                    normalizarTexto(
                        lista.name
                    );


                return (

                    displayName ===
                    objetivo

                    ||

                    name ===
                    objetivo

                );

            }

        );


    if (listaPermisos) {

        console.log(

            "Lista PermisosTI encontrada:",

            listaPermisos

        );


        return listaPermisos;

    }


    /* =====================================================
       MÉTODO 2
       BUSCAR MEDIANTE FILTER
    ===================================================== */

    try {

        const nombreSeguro =
            PERMISOS_LIST_NAME
                .replace(
                    /'/g,
                    "''"
                );


        const urlFiltro =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${sitio.id}/lists` +

            `?$select=id,name,displayName` +

            `&$filter=displayName eq '${nombreSeguro}'`;


        console.log(
            "Intentando búsqueda directa de PermisosTI:",
            urlFiltro
        );


        const resultadoFiltro =
            await graphFetch(
                urlFiltro
            );


        if (

            resultadoFiltro &&

            Array.isArray(
                resultadoFiltro.value
            ) &&

            resultadoFiltro.value.length

        ) {

            listaPermisos =
                resultadoFiltro.value[0];


            console.log(

                "PermisosTI encontrada mediante filtro:",

                listaPermisos

            );


            return listaPermisos;

        }

    }

    catch (errorFiltro) {

        console.warn(

            "La búsqueda filtrada de PermisosTI no estuvo disponible:",

            errorFiltro

        );

    }


    /* =====================================================
       ERROR DETALLADO
    ===================================================== */

    const nombresDisponibles =

        listas.length

            ? listas

                .map(

                    lista => {

                        const nombre =
                            lista.displayName ||
                            lista.name ||
                            "Sin nombre";


                        return `${nombre}`;

                    }

                )

                .join(", ")

            : "ninguna";


    const error =
        new Error(

            `No se encontró la lista "${PERMISOS_LIST_NAME}". ` +

            `Listas disponibles: ${nombresDisponibles}`

        );


    error.codigo =
        "LISTA_PERMISOS_NO_ENCONTRADA";


    error.sitioId =
        sitio.id;


    error.listasDisponibles =
        listas;


    console.error(

        "PERMISOSTI NO DISPONIBLE EN GRAPH",

        {

            listaBuscada:
                PERMISOS_LIST_NAME,

            sitio:
                sitio,

            listas:
                listas

        }

    );


    throw error;

}


    /* =========================================================
       OBTENER COLUMNAS
    ========================================================= */

    async function obtenerColumnasPermisos() {

        const sitio =
            await obtenerSitioSharePoint();


        const lista =
            await obtenerListaPermisos();


        const url =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${sitio.id}/lists/${lista.id}/columns` +

            `?$select=name,displayName,hidden,readOnly&$top=200`;


        const resultado =
            await graphFetch(
                url
            );


        return resultado.value || [];

    }


    /* =========================================================
       OBTENER MAPA DE CAMPOS
    ========================================================= */

    async function obtenerMapaCampos() {

        if (
            mapaCampos
        ) {

            return mapaCampos;

        }


        const columnas =
            await obtenerColumnasPermisos();


        mapaCampos =
            {};


        /*
         * COLUMNAS GENERALES
         */

        mapaCampos.Title =
            buscarColumna(

                columnas,

                [

                    "Title",

                    "Título"

                ]

            );


        mapaCampos.UsuarioCorreo =
            buscarColumna(

                columnas,

                [

                    "UsuarioCorreo"

                ]

            );


        mapaCampos.NombreUsuario =
            buscarColumna(

                columnas,

                [

                    "NombreUsuario"

                ]

            );


        mapaCampos.Grupo =
            buscarColumna(

                columnas,

                [

                    "Grupo"

                ]

            );


        if (!mapaCampos.Title) {

            throw new Error(
                "No se encontró la columna Title/Título."
            );

        }


        if (!mapaCampos.UsuarioCorreo) {

            throw new Error(
                "No se encontró la columna UsuarioCorreo."
            );

        }


        if (!mapaCampos.NombreUsuario) {

            throw new Error(
                "No se encontró la columna NombreUsuario."
            );

        }


        if (!mapaCampos.Grupo) {

            throw new Error(
                "No se encontró la columna Grupo."
            );

        }


        /*
         * COLUMNAS DE PERMISOS
         */

        MODULOS_PERMISOS.forEach(

            modulo => {

                mapaCampos[modulo] =
                    buscarColumna(

                        columnas,

                        [

                            modulo

                        ]

                    );


                if (
                    !mapaCampos[modulo]
                ) {

                    throw new Error(

                        `No se encontró la columna "${modulo}" en PermisosTI.`

                    );

                }

            }

        );


        console.log(

            "Mapa de campos PermisosTI:",

            mapaCampos

        );


        return mapaCampos;

    }


    /* =========================================================
       BUSCAR COLUMNA
    ========================================================= */

    function buscarColumna(
        columnas,
        nombres
    ) {

        const objetivos =
            nombres.map(
                normalizarTexto
            );


        let columna =
            columnas.find(

                item =>

                    objetivos.includes(

                        normalizarTexto(
                            item.displayName
                        )

                    )

            );


        if (!columna) {

            columna =
                columnas.find(

                    item =>

                        objetivos.includes(

                            normalizarTexto(
                                item.name
                            )

                        )

                );

        }


        return columna
            ? columna.name
            : null;

    }


    /* =========================================================
       OBTENER TODOS LOS REGISTROS
    ========================================================= */

    async function obtenerTodosLosItems() {

        const sitio =
            await obtenerSitioSharePoint();


        const lista =
            await obtenerListaPermisos();


        let url =

            `https://graph.microsoft.com/v1.0/sites/` +

            `${sitio.id}/lists/${lista.id}/items` +

            `?$expand=fields&$top=500`;


        const items =
            [];


        while (url) {

            const resultado =
                await graphFetch(
                    url
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


            url =
                resultado[
                    "@odata.nextLink"
                ] ||
                null;

        }


        console.log(
            "Registros obtenidos desde PermisosTI:",
            items
        );


        return items;

    }


    /* =========================================================
       CARGAR PERMISOS
    ========================================================= */

    async function cargarPermisos() {

        if (
            tablaPermisos
        ) {

            tablaPermisos.innerHTML =

                `
                <div class="tabla-vacia">

                    <span>
                        Cargando configuraciones...
                    </span>

                </div>
                `;

        }


        try {

            registrosPermisos =
                await obtenerTodosLosItems();


            renderizarTablaPermisos(
                registrosPermisos
            );


            console.log(

                "Configuraciones cargadas:",

                registrosPermisos.length

            );


        }

        catch (error) {

            console.error(

                "Error cargando PermisosTI:",

                error

            );


            if (
                tablaPermisos
            ) {

                tablaPermisos.innerHTML =

                    `
                    <div class="tabla-vacia">

                        <span>
                            No se pudieron cargar las configuraciones.
                        </span>

                    </div>
                    `;

            }


            throw error;

        }

    }


    /* =========================================================
       RENDERIZAR TABLA
    ========================================================= */

    function renderizarTablaPermisos(
        registros
    ) {

        if (
            !tablaPermisos
        ) {

            return;

        }


        if (
            !registros ||
            !registros.length
        ) {

            tablaPermisos.innerHTML =

                `
                <div class="tabla-vacia">

                    <span>
                        No hay usuarios configurados.
                    </span>

                </div>
                `;


            return;

        }


        let filas =
            "";


        registros.forEach(

            registro => {

                const correo =
                    obtenerCampo(
                        registro,
                        "UsuarioCorreo"
                    );


                const nombre =
                    obtenerCampo(
                        registro,
                        "NombreUsuario"
                    ) ||

                    obtenerCampo(
                        registro,
                        "Title"
                    ) ||

                    correo ||

                    "Sin nombre";


                const grupo =
                    String(

                        obtenerCampo(
                            registro,
                            "Grupo"
                        ) ||

                        GRUPO_NORMAL

                    ).trim();


                const cantidad =
                    contarPermisos(
                        registro
                    );


                const esTI =
                    normalizarTexto(
                        grupo
                    ) ===
                    normalizarTexto(
                        GRUPO_TI
                    );


                const completo =
                    esTI ||

                    cantidad ===
                        MODULOS_PERMISOS.length;


                filas +=

                    `
                    <tr>

                        <td>

                            <div class="permission-user-cell">

                                <div class="permission-user-avatar">
                                    👤
                                </div>

                                <div class="permission-user-info">

                                    <strong>
                                        ${escaparHTML(
                                            nombre
                                        )}
                                    </strong>

                                    <span>
                                        ${escaparHTML(
                                            correo
                                        )}
                                    </span>

                                </div>

                            </div>

                        </td>


                        <td>

                            <span
                                class="
                                    badge
                                    ${
                                        esTI
                                            ? "badge-ti"
                                            : "badge-normal"
                                    }
                                "
                            >
                                ${escaparHTML(
                                    grupo
                                )}
                            </span>

                        </td>


                        <td>

                            <span
                                class="
                                    access-summary
                                    ${
                                        completo
                                            ? "full"
                                            : "partial"
                                    }
                                "
                            >

                                ${
                                    esTI

                                        ? "Acceso total"

                                        : `${cantidad}/${MODULOS_PERMISOS.length}`

                                }

                            </span>

                        </td>


                        <td>

                            <button

                                type="button"

                                class="table-action-btn"

                                data-editar-permiso="${escaparHTML(
                                    registro.id
                                )}"

                            >

                                Editar

                            </button>

                        </td>

                    </tr>
                    `;

            }

        );


        tablaPermisos.innerHTML =

            `
            <div class="permissions-table-container">

                <table class="permissions-table">

                    <thead>

                        <tr>

                            <th>
                                Usuario
                            </th>

                            <th>
                                Grupo
                            </th>

                            <th>
                                Acceso
                            </th>

                            <th>
                                Acción
                            </th>

                        </tr>

                    </thead>

                    <tbody>

                        ${filas}

                    </tbody>

                </table>

            </div>
            `;


        tablaPermisos
            .querySelectorAll(
                "[data-editar-permiso]"
            )
            .forEach(

                boton => {

                    boton.addEventListener(

                        "click",

                        function () {

                            editarRegistro(

                                this.getAttribute(
                                    "data-editar-permiso"
                                )

                            );

                        }

                    );

                }

            );

    }


    /* =========================================================
       BUSCAR REGISTRO POR CORREO
    ========================================================= */

    function buscarRegistroPorCorreo(
        correo
    ) {

        const objetivo =
            normalizarCorreo(
                correo
            );


        if (!objetivo) {

            return null;

        }


        return registrosPermisos.find(

            registro =>

                normalizarCorreo(

                    obtenerCampo(
                        registro,
                        "UsuarioCorreo"
                    )

                ) === objetivo

        ) || null;

    }


    /* =========================================================
       CARGAR REGISTRO
    ========================================================= */

    function cargarRegistroEnFormulario(
        registro
    ) {

        const grupo =
            obtenerCampo(
                registro,
                "Grupo"
            );


        if (grupoUsuario) {

            grupoUsuario.value =

                normalizarTexto(
                    grupo
                ) ===
                normalizarTexto(
                    GRUPO_TI
                )

                    ? GRUPO_TI

                    : GRUPO_NORMAL;

        }


        MODULOS_PERMISOS.forEach(

            modulo => {

                const checkbox =
                    document.querySelector(

                        `.permiso-modulo[data-modulo="${modulo}"]`

                    );


                if (!checkbox) {

                    return;

                }


                checkbox.checked =
                    valorBooleano(

                        obtenerCampo(
                            registro,
                            modulo
                        )

                    );

            }

        );


        cambiarGrupo();

    }


    /* =========================================================
       PREPARAR NUEVO USUARIO
    ========================================================= */

    function prepararNuevoUsuario() {

        if (grupoUsuario) {

            grupoUsuario.value =
                GRUPO_NORMAL;

        }


        MODULOS_PERMISOS.forEach(

            modulo => {

                const checkbox =
                    document.querySelector(

                        `.permiso-modulo[data-modulo="${modulo}"]`

                    );


                if (!checkbox) {

                    return;

                }


                checkbox.checked =
                    false;


                checkbox.disabled =
                    false;

            }

        );

    }


    /* =========================================================
       CAMBIAR GRUPO
    ========================================================= */

    function cambiarGrupo() {

        const grupo =
            grupoUsuario
                ? grupoUsuario.value
                : GRUPO_NORMAL;


        const esTI =
            normalizarTexto(
                grupo
            ) ===
            normalizarTexto(
                GRUPO_TI
            );


        document
            .querySelectorAll(
                ".permiso-modulo"
            )
            .forEach(

                checkbox => {

                    if (esTI) {

                        checkbox.checked =
                            true;

                        checkbox.disabled =
                            true;

                    }

                    else {

                        checkbox.disabled =
                            false;

                    }

                }

            );

    }


    /* =========================================================
       GUARDAR PERMISOS
    ========================================================= */

    async function guardarPermisos() {

        if (!usuarioSeleccionado) {

            alert(
                "Primero selecciona un usuario."
            );


            return;

        }


        /*
         * Verificación adicional de seguridad.
         */

        const autorizado =
            await validarAdministradorTI();


        if (!autorizado) {

            return;

        }


        const grupo =
            grupoUsuario
                ? grupoUsuario.value
                : GRUPO_NORMAL;


        const permisos =
            {};


        MODULOS_PERMISOS.forEach(

            modulo => {

                const checkbox =
                    document.querySelector(

                        `.permiso-modulo[data-modulo="${modulo}"]`

                    );


                permisos[
                    modulo
                ] =

                    checkbox
                        ? checkbox.checked
                        : false;

            }

        );


        /*
         * TI siempre tiene todos los módulos.
         */

        if (
            normalizarTexto(
                grupo
            ) ===
            normalizarTexto(
                GRUPO_TI
            )
        ) {

            MODULOS_PERMISOS.forEach(

                modulo => {

                    permisos[
                        modulo
                    ] =
                        true;

                }

            );

        }


        try {

            if (btnGuardarPermisos) {

                btnGuardarPermisos.disabled =
                    true;

                btnGuardarPermisos.textContent =
                    "Guardando...";

            }


            const sitio =
                await obtenerSitioSharePoint();


            const lista =
                await obtenerListaPermisos();


            const campos =
                await obtenerMapaCampos();


            const existente =
                buscarRegistroPorCorreo(

                    usuarioSeleccionado.correo

                );


            const fields =
                {};


            /*
             * NOMBRE
             */

            fields[
                campos.Title
            ] =

                usuarioSeleccionado.nombre ||
                usuarioSeleccionado.correo;


            /*
             * CORREO
             */

            fields[
                campos.UsuarioCorreo
            ] =

                usuarioSeleccionado.correo;


            /*
             * NOMBRE USUARIO
             */

            fields[
                campos.NombreUsuario
            ] =

                usuarioSeleccionado.nombre ||
                usuarioSeleccionado.correo;


            /*
             * GRUPO
             */

            fields[
                campos.Grupo
            ] =
                grupo;


            /*
             * PERMISOS
             */

            MODULOS_PERMISOS.forEach(

                modulo => {

                    fields[
                        campos[modulo]
                    ] =

                        Boolean(

                            permisos[
                                modulo
                            ]

                        );

                }

            );


            /*
             * ACTUALIZAR REGISTRO EXISTENTE
             */

            if (existente) {

                const url =

                    `https://graph.microsoft.com/v1.0/sites/` +

                    `${sitio.id}/lists/${lista.id}/items/` +

                    `${existente.id}/fields`;


                await graphFetch(

                    url,

                    {

                        method:
                            "PATCH",

                        body:
                            JSON.stringify(
                                fields
                            )

                    }

                );


                mostrarMensaje(

                    "Permisos actualizados correctamente.",

                    "success"

                );

            }


            /*
             * CREAR REGISTRO NUEVO
             */

            else {

                const url =

                    `https://graph.microsoft.com/v1.0/sites/` +

                    `${sitio.id}/lists/${lista.id}/items`;


                await graphFetch(

                    url,

                    {

                        method:
                            "POST",

                        body:
                            JSON.stringify({

                                fields:
                                    fields

                            })

                    }

                );


                mostrarMensaje(

                    "Permisos guardados correctamente.",

                    "success"

                );

            }


            /*
             * Recargar configuraciones.
             */

            await cargarPermisos();


            /*
             * Volver a cargar usuario.
             */

            const actualizado =
                buscarRegistroPorCorreo(

                    usuarioSeleccionado.correo

                );


            if (actualizado) {

                cargarRegistroEnFormulario(
                    actualizado
                );

            }

        }

        catch (error) {

            console.error(

                "Error guardando permisos:",

                error

            );


            mostrarMensaje(

                "No se pudieron guardar los permisos. " +

                obtenerMensajeError(
                    error
                ),

                "error"

            );

        }

        finally {

            if (btnGuardarPermisos) {

                btnGuardarPermisos.disabled =
                    false;

                btnGuardarPermisos.textContent =
                    "Guardar permisos";

            }

        }

    }


    /* =========================================================
       EDITAR REGISTRO
    ========================================================= */

    async function editarRegistro(
        itemId
    ) {

        const registro =
            registrosPermisos.find(

                item =>

                    String(
                        item.id
                    ) ===
                    String(
                        itemId
                    )

            );


        if (!registro) {

            alert(
                "No se encontró el registro."
            );


            return;

        }


        const correo =
            normalizarCorreo(

                obtenerCampo(
                    registro,
                    "UsuarioCorreo"
                )

            );


        const usuario =
            usuariosGraph.find(

                item =>

                    normalizarCorreo(

                        item.mail ||

                        item.userPrincipalName ||

                        ""

                    ) === correo

            );


        if (!usuario) {

            alert(

                "El usuario ya no se encuentra en Microsoft Graph."

            );


            return;

        }


        if (selectorUsuario) {

            selectorUsuario.value =
                usuario.id;

        }


        seleccionarUsuario(
            usuario.id
        );

    }


    /* =========================================================
       VALIDAR ADMIN TI
    ========================================================= */

    async function validarAdministradorTI() {

        const usuarioActualCorreo =
            obtenerCorreoCuentaActual();


        console.log(
            "Validando administrador TI:",
            usuarioActualCorreo
        );


        /*
         * TI INICIAL
         */

        const esTIInicial =

            USUARIOS_TI_INICIALES
                .map(
                    normalizarCorreo
                )
                .includes(
                    normalizarCorreo(
                        usuarioActualCorreo
                    )
                );


        if (
            esTIInicial
        ) {

            console.log(
                "Administrador TI autorizado mediante usuario inicial."
            );


            return true;

        }


        /*
         * PERMISOSTI
         */

        const registro =
            buscarRegistroPorCorreo(
                usuarioActualCorreo
            );


        if (
            registro &&
            normalizarTexto(

                obtenerCampo(
                    registro,
                    "Grupo"
                )

            ) ===
                normalizarTexto(
                    GRUPO_TI
                )
        ) {

            console.log(
                "Administrador TI autorizado mediante PermisosTI."
            );


            return true;

        }


        console.warn(
            "Usuario no autorizado como administrador TI."
        );


        mostrarAccesoDenegado();


        return false;

    }


    /* =========================================================
       CORREO CUENTA ACTUAL
    ========================================================= */

    function obtenerCorreoCuentaActual() {

        /*
         * Primero usamos auth.js.
         */

        if (
            window.alferzaAuth &&
            typeof window.alferzaAuth.obtenerCorreo ===
                "function"
        ) {

            const correoAuth =
                window.alferzaAuth.obtenerCorreo();


            if (correoAuth) {

                return normalizarCorreo(
                    correoAuth
                );

            }

        }


        /*
         * Respaldo.
         */

        if (!cuentaActual) {

            return "";

        }


        return normalizarCorreo(

            cuentaActual.username ||

            cuentaActual.mail ||

            cuentaActual.userPrincipalName ||

            cuentaActual.idTokenClaims?.preferred_username ||

            cuentaActual.idTokenClaims?.email ||

            ""

        );

    }


    /* =========================================================
       MOSTRAR ACCESO DENEGADO
    ========================================================= */

    function mostrarAccesoDenegado() {

        const contenido =
            document.querySelector(
                ".content"
            );


        if (!contenido) {

            return;

        }


        contenido.innerHTML =

            `
            <section class="page-header">

                <div>

                    <h1>
                        Permisos
                    </h1>

                    <p>
                        Administración de accesos al sistema
                    </p>

                </div>

            </section>


            <div class="permisos-card">

                <div class="tabla-vacia">

                    <span>
                        No tienes permisos para administrar esta sección.
                    </span>

                </div>

            </div>
            `;

    }


    /* =========================================================
       VALOR BOOLEANO
    ========================================================= */

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


    /* =========================================================
       OBTENER CAMPO
    ========================================================= */

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


    /* =========================================================
       CONTAR PERMISOS
    ========================================================= */

    function contarPermisos(
        registro
    ) {

        const grupo =
            normalizarTexto(

                obtenerCampo(
                    registro,
                    "Grupo"
                )

            );


        if (
            grupo ===
            normalizarTexto(
                GRUPO_TI
            )
        ) {

            return MODULOS_PERMISOS.length;

        }


        let cantidad =
            0;


        MODULOS_PERMISOS.forEach(

            modulo => {

                if (

                    valorBooleano(

                        obtenerCampo(
                            registro,
                            modulo
                        )

                    )

                ) {

                    cantidad++;

                }

            }

        );


        return cantidad;

    }


    /* =========================================================
       NORMALIZAR TEXTO
    ========================================================= */

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


    /* =========================================================
       NORMALIZAR CORREO
    ========================================================= */

    function normalizarCorreo(
        correo
    ) {

        return String(
            correo || ""
        )
            .trim()
            .toLowerCase();

    }


    /* =========================================================
       ESCAPAR HTML
    ========================================================= */

    function escaparHTML(
        valor
    ) {

        if (
            valor === null ||
            valor === undefined
        ) {

            return "";

        }


        return String(
            valor
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
       MENSAJE
    ========================================================= */

    function mostrarMensaje(
        mensaje,
        tipo = "error"
    ) {

        const contenido =
            document.querySelector(
                ".content"
            );


        if (!contenido) {

            return;

        }


        const anterior =
            contenido.querySelector(
                ".mensaje-permisos"
            );


        if (anterior) {

            anterior.remove();

        }


        const aviso =
            document.createElement(
                "div"
            );


        aviso.className =
            `permission-message ${tipo} mensaje-permisos`;


        aviso.textContent =
            mensaje;


        contenido.prepend(
            aviso
        );


        setTimeout(

            function () {

                if (
                    aviso &&
                    aviso.parentNode
                ) {

                    aviso.remove();

                }

            },

            5000

        );

    }


    /* =========================================================
       ERROR INICIAL
    ========================================================= */

    function mostrarErrorInicial(
        mensaje
    ) {

        const usuario =
            document.getElementById(
                "userInfo"
            );


        if (usuario) {

            usuario.textContent =
                mensaje;

        }

    }


    /* =========================================================
       ERROR
    ========================================================= */

    function obtenerMensajeError(
        error
    ) {

        if (
            !error
        ) {

            return "Error desconocido.";

        }


        const mensaje =
            String(
                error.message ||
                error
            );


        if (
            mensaje.includes(
                "Graph 401"
            )
        ) {

            return (

                "La sesión de Microsoft no es válida. Cierra sesión y vuelve a iniciar sesión."

            );

        }


        if (
            mensaje.includes(
                "Graph 403"
            )
        ) {

            return (

                "Microsoft Graph rechazó la operación. Revisa los permisos concedidos a ALFERZA Live Office."

            );

        }


        if (
            error.codigo ===
            "LISTA_PERMISOS_NO_ENCONTRADA"
        ) {

            return (

                "Microsoft Graph no está mostrando la lista PermisosTI en el sitio configurado."

            );

        }


        return mensaje;

    }


    /* =========================================================
       EXPONER FUNCIONES
    ========================================================= */

    window.seleccionarUsuario =
        seleccionarUsuario;


    window.buscarUsuario =
        buscarUsuario;


    window.guardarPermisos =
        guardarPermisos;


    window.editarRegistro =
        editarRegistro;


    window.validarAdministradorTI =
        validarAdministradorTI;


})();