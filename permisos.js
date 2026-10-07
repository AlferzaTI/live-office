
(function () {

    /* =========================================================
       ALFERZA LIVE OFFICE
       GESTIÓN DE PERMISOS
    ========================================================= */


    /* =========================================================
       CONFIGURACIÓN SHAREPOINT
    ========================================================= */

    const SHAREPOINT_HOST =
        "alferzaholding-my.sharepoint.com";


    const SHAREPOINT_SITE_PATH =
        "/personal/soporte1_alferza_pe";


    const SHAREPOINT_BASE_URL =
        `https://${SHAREPOINT_HOST}${SHAREPOINT_SITE_PATH}`;


    const PERMISOS_LIST_NAME =
        "PermisosTI";


    const PERMISOS_LIST_ID =
        "9C313E91-5655-44BF-975C-4BEB9D56C2C1";


    const PERMISOS_LIST_URL =
        `${SHAREPOINT_SITE_PATH}/Lists/${PERMISOS_LIST_NAME}`;


    /* =========================================================
       SCOPES GRAPH
       SOLO PARA MICROSOFT GRAPH
    ========================================================= */

    const SCOPES_GRAPH = [

        "User.Read",

        "Sites.ReadWrite.All"

    ];


    /* =========================================================
       SCOPES SHAREPOINT
       IMPORTANTE:
       ESTE TOKEN ES DIFERENTE AL TOKEN DE GRAPH
    ========================================================= */

    const SCOPES_SHAREPOINT_READ = [

        `https://${SHAREPOINT_HOST}/AllSites.Read`

    ];


    const SCOPES_SHAREPOINT_WRITE = [

        `https://${SHAREPOINT_HOST}/AllSites.Write`

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
       SOLO CUENTA DE EMERGENCIA / ADMINISTRACIÓN INICIAL
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
       USA LA MISMA INSTANCIA DE AUTH.JS
    ========================================================= */

    const msalInstance =
        window.alferzaMsal;


    /* =========================================================
       VARIABLES
    ========================================================= */

    let cuentaActual =
        null;


    let usuariosGraph =
        [];


    let registrosPermisos =
        [];


    let usuarioSeleccionado =
        null;


    let mapaCampos =
        null;


    let tipoEntidadLista =
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
               ESPERAR AUTH.JS
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
               CARGAR USUARIOS DESDE GRAPH
            ================================================= */

            await cargarUsuariosGraph();


            /* =================================================
               CARGAR PERMISOS DESDE SHAREPOINT REST
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
                 * El administrador inicial puede continuar.
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

        if (btnBuscarUsuario) {

            btnBuscarUsuario.addEventListener(

                "click",

                buscarUsuario

            );

        }


        if (correoUsuario) {

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


        if (selectorUsuario) {

            selectorUsuario.addEventListener(

                "change",

                function () {

                    seleccionarUsuario(
                        this.value
                    );

                }

            );

        }


        if (grupoUsuario) {

            grupoUsuario.addEventListener(

                "change",

                cambiarGrupo

            );

        }


        if (btnGuardarPermisos) {

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
       TOKEN GRAPH
    ========================================================= */

    async function obtenerTokenGraph() {

        if (

            window.alferzaAuth &&

            typeof window.alferzaAuth.obtenerToken ===
                "function"

        ) {

            return await window.alferzaAuth.obtenerToken(

                SCOPES_GRAPH

            );

        }


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


        const resultado =

            await msalInstance.acquireTokenSilent({

                scopes:
                    SCOPES_GRAPH,

                account:
                    cuenta

            });


        return resultado.accessToken;

    }


    /* =========================================================
       TOKEN SHAREPOINT
       USA LA MISMA SESIÓN MSAL DE AUTH.JS
    ========================================================= */

    async function obtenerTokenSharePoint(
        escritura = false
    ) {

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


        const scopes =

            escritura

                ? SCOPES_SHAREPOINT_WRITE

                : SCOPES_SHAREPOINT_READ;


        console.log(

            "Solicitando token SharePoint:",

            scopes

        );


        try {

            const resultado =

                await msalInstance.acquireTokenSilent({

                    scopes:
                        scopes,

                    account:
                        cuenta

                });


            if (

                !resultado ||

                !resultado.accessToken

            ) {

                throw new Error(

                    "Microsoft no devolvió un token de SharePoint."

                );

            }


            return resultado.accessToken;

        }

        catch (error) {

            console.error(

                "Error obteniendo token SharePoint:",

                error

            );


            const nuevoError =
                new Error(

                    "No se pudo obtener el token de SharePoint. " +

                    "Verifica que la aplicación tenga permisos " +

                    "SharePoint AllSites.Read/AllSites.Write."

                );


            nuevoError.codigo =
                "SHAREPOINT_TOKEN_ERROR";


            nuevoError.errorOriginal =
                error;


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


    /* =========================================================
       SHAREPOINT REST FETCH
    ========================================================= */

    async function sharePointFetch(

        url,

        opciones = {},

        escritura = false

    ) {

        const token =

            await obtenerTokenSharePoint(

                escritura

            );


        const headers = {

            Authorization:
                `Bearer ${token}`,

            Accept:
                "application/json;odata=nometadata",

            "Content-Type":
                "application/json;odata=nometadata",

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


        const texto =
            await respuesta.text();


        if (!respuesta.ok) {

            console.error(

                "SharePoint REST error:",

                {

                    status:
                        respuesta.status,

                    url:
                        url,

                    response:
                        texto

                }

            );


            let detalle =
                texto;


            try {

                const json =
                    JSON.parse(
                        texto
                    );


                detalle =

                    json.error?.message?.value ||

                    json.error?.message ||

                    texto;

            }

            catch {

                // La respuesta no era JSON.

            }


            const error =
                new Error(

                    `SharePoint ${respuesta.status}: ${detalle}`

                );


            error.status =
                respuesta.status;


            error.response =
                texto;


            throw error;

        }


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

        if (selectorUsuario) {

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

                            ).toLowerCase();


                        const nombreB =

                            String(

                                b.displayName ||

                                ""

                            ).toLowerCase();


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

        if (!selectorUsuario) {

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


        if (!usuarios.length) {

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

        if (panelPermisos) {

            panelPermisos.style.display =
                "block";

        }

    }


    /* =========================================================
       OCULTAR PANEL
    ========================================================= */

    function ocultarPanel() {

        if (panelPermisos) {

            panelPermisos.style.display =
                "none";

        }

    }


    /* =========================================================
       OBTENER LISTA MEDIANTE SHAREPOINT REST
    ========================================================= */

    async function obtenerListaSharePoint() {

        const url =

            `${SHAREPOINT_BASE_URL}` +

            `/_api/web/GetList(@listUrl)` +

            `?$select=Id,Title,ListItemEntityTypeFullName` +

            `&@listUrl='${encodeURIComponent(

                PERMISOS_LIST_URL

            )}'`;


        console.log(

            "Obteniendo información de PermisosTI mediante SharePoint REST:",

            url

        );


        const resultado =

            await sharePointFetch(

                url

            );


        console.log(

            "Información PermisosTI:",

            resultado

        );


        return resultado;

    }


    /* =========================================================
       OBTENER ITEMS DE PERMISOSTI
       RenderListDataAsStream
    ========================================================= */

    async function obtenerTodosLosItems() {

        /*
         * IMPORTANTE:
         * No utilizamos Graph.
         *
         * SharePoint está usando exactamente este endpoint
         * para cargar la lista PermisosTI.
         */

        const url =

            `${SHAREPOINT_BASE_URL}` +

            `/_api/web/GetList(@listUrl)/RenderListDataAsStream` +

            `?@listUrl='${encodeURIComponent(

                PERMISOS_LIST_URL

            )}'`;


        console.log(

            "Consultando PermisosTI mediante RenderListDataAsStream:",

            url

        );


        /*
         * CAML SIN FILTRO DE FECHA.
         *
         * Así obtenemos los 4 registros de la lista.
         */

        const viewXml =

            `<View Scope="RecursiveAll">

                <Query>

                    <OrderBy>

                        <FieldRef Name="ID" Ascending="TRUE" />

                    </OrderBy>

                </Query>

                <ViewFields>

                    <FieldRef Name="ID" />

                    <FieldRef Name="Title" />

                    <FieldRef Name="UsuarioCorreo" />

                    <FieldRef Name="NombreUsuario" />

                    <FieldRef Name="Grupo" />

                    <FieldRef Name="Oficina" />

                    <FieldRef Name="Personal" />

                    <FieldRef Name="Reservas" />

                    <FieldRef Name="Salas" />

                    <FieldRef Name="Comunicados" />

                    <FieldRef Name="Seguridad" />

                    <FieldRef Name="Infraestructura" />

                    <FieldRef Name="Tickets" />

                    <FieldRef Name="Configuracion" />

                </ViewFields>

                <RowLimit Paged="TRUE">

                    100

                </RowLimit>

            </View>`;


        const body = {

            parameters: {

                __metadata: {

                    type:
                        "SP.RenderListDataParameters"

                },

                AddRequiredFields:
                    true,

                AllowMultipleValueFilterForTaxonomyFields:
                    true,

                FilterOutChannelFoldersInDefaultDocLib:
                    true,

                RequireFolderColoringFields:
                    true,

                DatesInUtc:
                    true,

                RenderOptions:
                    1183783,

                ViewXml:
                    viewXml

            }

        };


        console.log(

            "Body RenderListDataAsStream:",

            body

        );


        const resultado =

            await sharePointFetch(

                url,

                {

                    method:
                        "POST",

                    body:
                        JSON.stringify(
                            body
                        )

                },

                false

            );


        console.log(

            "Respuesta PermisosTI:",

            resultado

        );


        if (

            !resultado

        ) {

            throw new Error(

                "SharePoint no devolvió información de PermisosTI."

            );

        }


        const filas =

            resultado.ListData &&

            Array.isArray(
                resultado.ListData.Row
            )

                ? resultado.ListData.Row

                : [];


        console.log(

            "Filas encontradas en PermisosTI:",

            filas.length

        );


        console.log(

            "Registros PermisosTI:",

            filas

        );


        /*
         * Convertimos las filas de SharePoint
         * a la estructura que ya utiliza el resto
         * del código.
         */

        const items =

            filas.map(

                fila => {

                    const fields = {};


                    Object.keys(
                        fila
                    ).forEach(

                        clave => {

                            fields[clave] =
                                fila[clave];

                        }

                    );


                    return {

                        id:

                            fila.ID ||

                            fila.Id ||

                            fila.id ||

                            "",

                        fields:
                            fields,

                        raw:
                            fila

                    };

                }

            );


        return items;

    }


    /* =========================================================
       CARGAR PERMISOS
    ========================================================= */

    async function cargarPermisos() {

        if (tablaPermisos) {

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


            if (tablaPermisos) {

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

        if (!tablaPermisos) {

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
       OBTENER FORM DIGEST
       NECESARIO PARA OPERACIONES POST DE ESCRITURA
    ========================================================= */

    async function obtenerRequestDigest() {

        const url =

            `${SHAREPOINT_BASE_URL}` +

            `/_api/contextinfo`;


        const resultado =

            await sharePointFetch(

                url,

                {

                    method:
                        "POST",

                    body:
                        "{}"

                },

                true

            );


        const digest =

            resultado

                ?.d

                ?.GetContextWebInformation

                ?.FormDigestValue;


        if (!digest) {

            throw new Error(

                "SharePoint no devolvió X-RequestDigest."

            );

        }


        return digest;

    }


    /* =========================================================
       OBTENER TIPO DE ITEM DE LA LISTA
    ========================================================= */

    async function obtenerTipoEntidadLista() {

        if (tipoEntidadLista) {

            return tipoEntidadLista;

        }


        const url =

            `${SHAREPOINT_BASE_URL}` +

            `/_api/web/GetList(@listUrl)` +

            `?$select=ListItemEntityTypeFullName` +

            `&@listUrl='${encodeURIComponent(

                PERMISOS_LIST_URL

            )}'`;


        const resultado =

            await sharePointFetch(

                url

            );


        tipoEntidadLista =

            resultado.ListItemEntityTypeFullName;


        if (!tipoEntidadLista) {

            throw new Error(

                "No se pudo obtener ListItemEntityTypeFullName de PermisosTI."

            );

        }


        console.log(

            "Tipo de entidad PermisosTI:",

            tipoEntidadLista

        );


        return tipoEntidadLista;

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
         * Verificación adicional.
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


        const permisos = {};


        MODULOS_PERMISOS.forEach(

            modulo => {

                const checkbox =

                    document.querySelector(

                        `.permiso-modulo[data-modulo="${modulo}"]`

                    );


                permisos[modulo] =

                    checkbox

                        ? checkbox.checked

                        : false;

            }

        );


        /*
         * TI siempre tiene todo.
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

                    permisos[modulo] =
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


            const existente =

                buscarRegistroPorCorreo(

                    usuarioSeleccionado.correo

                );


            const fields = {};


            /*
             * SharePoint REST utiliza los nombres internos.
             */

            fields.Title =

                usuarioSeleccionado.nombre ||

                usuarioSeleccionado.correo;


            fields.UsuarioCorreo =

                usuarioSeleccionado.correo;


            fields.NombreUsuario =

                usuarioSeleccionado.nombre ||

                usuarioSeleccionado.correo;


            fields.Grupo =
                grupo;


            MODULOS_PERMISOS.forEach(

                modulo => {

                    fields[modulo] =

                        Boolean(

                            permisos[modulo]

                        );

                }

            );


            const digest =

                await obtenerRequestDigest();


            /*
             * =================================================
             * ACTUALIZAR
             * =================================================
             */

            if (existente) {

                const itemId =
                    existente.id;


                const url =

                    `${SHAREPOINT_BASE_URL}` +

                    `/_api/web/GetList(@listUrl)/items(${itemId})` +

                    `?@listUrl='${encodeURIComponent(

                        PERMISOS_LIST_URL

                    )}'`;


                const headers = {

                    "X-RequestDigest":
                        digest,

                    "IF-MATCH":
                        "*",

                    "X-HTTP-Method":
                        "MERGE",

                    Accept:
                        "application/json;odata=verbose",

                    "Content-Type":
                        "application/json;odata=verbose"

                };


                const token =

                    await obtenerTokenSharePoint(

                        true

                    );


                const respuesta =

                    await fetch(

                        url,

                        {

                            method:
                                "POST",

                            headers: {

                                Authorization:
                                    `Bearer ${token}`,

                                ...headers

                            },

                            body:
                                JSON.stringify(
                                    fields
                                )

                        }

                    );


                const texto =

                    await respuesta.text();


                if (!respuesta.ok) {

                    throw new Error(

                        `SharePoint ${respuesta.status}: ${texto}`

                    );

                }


                mostrarMensaje(

                    "Permisos actualizados correctamente.",

                    "success"

                );

            }


            /*
             * =================================================
             * CREAR
             * =================================================
             */

            else {

                const tipoEntidad =

                    await obtenerTipoEntidadLista();


                const url =

                    `${SHAREPOINT_BASE_URL}` +

                    `/_api/web/GetList(@listUrl)/items` +

                    `?@listUrl='${encodeURIComponent(

                        PERMISOS_LIST_URL

                    )}'`;


                const token =

                    await obtenerTokenSharePoint(

                        true

                    );


                const body = {

                    __metadata: {

                        type:
                            tipoEntidad

                    },

                    ...fields

                };


                const respuesta =

                    await fetch(

                        url,

                        {

                            method:
                                "POST",

                            headers: {

                                Authorization:
                                    `Bearer ${token}`,

                                Accept:
                                    "application/json;odata=verbose",

                                "Content-Type":
                                    "application/json;odata=verbose",

                                "X-RequestDigest":
                                    digest

                            },

                            body:
                                JSON.stringify(
                                    body
                                )

                        }

                    );


                const texto =

                    await respuesta.text();


                if (!respuesta.ok) {

                    throw new Error(

                        `SharePoint ${respuesta.status}: ${texto}`

                    );

                }


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


        if (esTIInicial) {

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

            registro.raw ||

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


        return fields[clave];

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

        if (!error) {

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

            mensaje.includes(
                "SharePoint 401"
            )

        ) {

            return (

                "El token de SharePoint no es válido o la aplicación no tiene acceso a SharePoint."

            );

        }


        if (

            mensaje.includes(
                "SharePoint 403"
            )

        ) {

            return (

                "SharePoint rechazó el acceso a PermisosTI. Revisa los permisos AllSites.Read/AllSites.Write de la aplicación."

            );

        }


        if (

            error.codigo ===
            "SHAREPOINT_TOKEN_ERROR"

        ) {

            return (

                "No se pudo obtener el token de SharePoint. Revisa los permisos SharePoint de la aplicación."

            );

        }


        if (

            error.codigo ===
            "LISTA_PERMISOS_NO_ENCONTRADA"

        ) {

            return (

                "No se encontró la lista PermisosTI."

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
