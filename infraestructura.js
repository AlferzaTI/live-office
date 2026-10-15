
"use strict";

/* =========================================================
   ALFERZA LIVE OFFICE — CHATBOT TI
   Sin backend propio.
   Microsoft Graph + Microsoft Lists + PDF.js.
========================================================= */

(function () {
    const MSAL_CONFIG = {
        auth: {
            clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
            authority:
                "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",
            redirectUri: window.location.hostname === "localhost"
                ? "http://localhost:5500/blank.html"
                : "https://alferzati.github.io/live-office/blank.html"
        },
        cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: false
        }
    };

    const SCOPES_GRAPH = [
        "User.Read",
        "Sites.Read.All",
        "Sites.ReadWrite.All"
    ];

    const GRAPH = "https://graph.microsoft.com/v1.0";

    // Lista de conversaciones: sitio personal.
    const LIST_SITE_HOST = "alferzaholding-my.sharepoint.com";
    const LIST_SITE_PATH = "/personal/soporte1_alferza_pe";
    const LIST_NAME = "ConversacionesCHATBOT";

    // Documentos: sitio de TI.
    const DOCS_SITE_HOST = "alferzaholding.sharepoint.com";
    const DOCS_SITE_PATH = "/sites/TI";
    const DOCS_FOLDER = "Procedimientos-TI";

    const STORAGE_KEY = "alferza_chatbot_conversation_id";
    const HISTORY_PREFIX = "alferza_chat_history_";
    const CACHE_PREFIX = "alferza_chat_cache_";

    const MAX_CARACTERES = 1000;
    const MAX_ENTRADAS_CACHE = 100;
    const MAX_PDFS = 100;

    const elements = {
        form: document.getElementById("chatForm"),
        input: document.getElementById("messageInput"),
        messages: document.getElementById("messages"),
        sendButton: document.getElementById("sendButton"),
        clearButton: document.getElementById("clearChat"),
        typing: document.getElementById("typingIndicator"),
        characterCount: document.getElementById("characterCount"),
        suggestions: document.querySelectorAll(".suggestion-button")
    };

    if (
        !elements.form ||
        !elements.input ||
        !elements.messages ||
        !elements.sendButton ||
        !elements.clearButton ||
        !elements.typing ||
        !elements.characterCount
    ) {
        console.error("[ALFERZA CHATBOT] Faltan elementos HTML.");
        return;
    }

    /* =====================================================
       VARIABLES
    ===================================================== */

    let msalInstance = null;
    let cuentaActual = null;
    let tokenActual = null;
    let usuarioCorreo = "";

    let sitioLista = null;
    let listaConversaciones = null;
    let columnasLista = null;

    let sitioDocumentos = null;
    let driveDocumentos = null;
    let archivosPDF = [];
    let indiceDocumentos = [];
    let promesaIndexacion = null;

    let procesandoMensaje = false;
    let conversationId = obtenerConversationId();
    let historial = cargarHistorial();
    let cache = cargarCache();

    /* =====================================================
       UTILIDADES
    ===================================================== */

    function normalizarTexto(valor) {
        return String(valor || "")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase()
            .replace(/\s+/g, " ")
            .trim();
    }

    function generarId() {
        if (window.crypto?.randomUUID) {
            return window.crypto.randomUUID();
        }

        return "chat-" + Date.now() + "-" +
            Math.random().toString(36).slice(2, 12);
    }

    function obtenerConversationId() {
        try {
            let id = sessionStorage.getItem(STORAGE_KEY);

            if (!id) {
                id = generarId();
                sessionStorage.setItem(STORAGE_KEY, id);
            }

            return id;
        } catch (_) {
            return generarId();
        }
    }

    function claveHistorial() {
        return HISTORY_PREFIX + conversationId;
    }

    function claveCache() {
        return CACHE_PREFIX + conversationId;
    }

    function leerStorage(clave, defecto) {
        try {
            const valor = sessionStorage.getItem(clave);
            return valor ? JSON.parse(valor) : defecto;
        } catch (_) {
            return defecto;
        }
    }

    function guardarStorage(clave, valor) {
        try {
            sessionStorage.setItem(clave, JSON.stringify(valor));
        } catch (_) {
            console.warn("[ALFERZA] No se pudo guardar en la sesión.");
        }
    }

    function cargarHistorial() {
        const valor = leerStorage(claveHistorial(), []);
        return Array.isArray(valor) ? valor : [];
    }

    function guardarHistorial() {
        guardarStorage(claveHistorial(), historial);
    }

    function cargarCache() {
        const valor = leerStorage(claveCache(), {});
        return valor && typeof valor === "object" &&
            !Array.isArray(valor) ? valor : {};
    }

    function guardarCache() {
        guardarStorage(claveCache(), cache);
    }

    function obtenerCorreoCuenta(cuenta) {
        return String(
            cuenta?.username ||
            cuenta?.idTokenClaims?.preferred_username ||
            cuenta?.idTokenClaims?.email ||
            ""
        ).trim().toLowerCase();
    }

    function escaparOData(valor) {
        return String(valor || "").replace(/'/g, "''");
    }

    function normalizarNombreColumna(valor) {
        return normalizarTexto(valor)
            .replace(/[^a-z0-9]/g, "");
    }

    /* =====================================================
       MSAL Y AUTENTICACIÓN
    ===================================================== */

    function cargarScript(src) {
        return new Promise((resolve, reject) => {
            const existente = [...document.scripts].find(
                script => script.src === src
            );

            if (existente) {
                if (existente.dataset.alferzaLoaded === "true") {
                    resolve();
                    return;
                }

                existente.addEventListener("load", resolve, { once: true });
                existente.addEventListener(
                    "error",
                    () => reject(new Error("No se pudo cargar un script.")),
                    { once: true }
                );
                return;
            }

            const script = document.createElement("script");
            script.src = src;
            script.async = true;

            script.onload = () => {
                script.dataset.alferzaLoaded = "true";
                resolve();
            };

            script.onerror = () => {
                reject(new Error("No se pudo cargar una biblioteca externa."));
            };

            document.head.appendChild(script);
        });
    }

    async function inicializarMSAL() {
        if (typeof msal === "undefined") {
            await cargarScript(
                "https://alcdn.msauth.net/browser/2.38.1/js/msal-browser.min.js"
            );
        }

        msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);

        if (typeof msalInstance.initialize === "function") {
            await msalInstance.initialize();
        }

        if (typeof msalInstance.handleRedirectPromise === "function") {
            const resultado = await msalInstance.handleRedirectPromise();

            if (resultado?.account) {
                cuentaActual = resultado.account;
            }
        }

        if (!cuentaActual) {
            const cuentas = msalInstance.getAllAccounts();

            if (cuentas.length) {
                cuentaActual = cuentas[0];
            }
        }

        if (cuentaActual) {
            msalInstance.setActiveAccount?.(cuentaActual);
        }

        return Boolean(cuentaActual);
    }

    async function obtenerTokenSilencioso() {
        if (!msalInstance) {
            await inicializarMSAL();
        }

        if (!cuentaActual) {
            const cuentas = msalInstance.getAllAccounts();

            if (cuentas.length) {
                cuentaActual = cuentas[0];
            }
        }

        if (!cuentaActual) {
            throw new Error(
                "No hay una cuenta Microsoft autenticada. Inicia sesión en ALFERZA."
            );
        }

        try {
            const respuesta = await msalInstance.acquireTokenSilent({
                scopes: SCOPES_GRAPH,
                account: cuentaActual
            });

            tokenActual = respuesta.accessToken;
            cuentaActual = respuesta.account || cuentaActual;
            usuarioCorreo = obtenerCorreoCuenta(cuentaActual);

            if (!usuarioCorreo) {
                throw new Error(
                    "No se pudo identificar el correo de la cuenta Microsoft."
                );
            }

            return tokenActual;
        } catch (error) {
            tokenActual = null;
            throw error;
        }
    }

    async function iniciarSesionInteractiva() {
        if (!msalInstance) {
            await inicializarMSAL();
        }

        const cuentas = msalInstance.getAllAccounts();

        const respuesta = cuentas.length
            ? await msalInstance.acquireTokenPopup({
                scopes: SCOPES_GRAPH,
                account: cuentas[0]
            })
            : await msalInstance.loginPopup({
                scopes: SCOPES_GRAPH
            });

        cuentaActual = respuesta.account;

        if (cuentaActual) {
            msalInstance.setActiveAccount?.(cuentaActual);
        }

        tokenActual = respuesta.accessToken || null;
        usuarioCorreo = obtenerCorreoCuenta(cuentaActual);

        if (!tokenActual) {
            await obtenerTokenSilencioso();
        }

        return tokenActual;
    }

    /* =====================================================
       MICROSOFT GRAPH
    ===================================================== */

    async function graphFetch(url, opciones = {}, intento = 0) {
        if (!tokenActual) {
            await obtenerTokenSilencioso();
        }

        const { timeoutMs = 30000, ...fetchOptions } = opciones;
        const controlador = new AbortController();

        const temporizador = setTimeout(
            () => controlador.abort(),
            timeoutMs
        );

        try {
            const respuesta = await fetch(url, {
                ...fetchOptions,
                signal: controlador.signal,
                headers: {
                    Authorization: `Bearer ${tokenActual}`,
                    ...fetchOptions.headers
                }
            });

            if (
                (respuesta.status === 429 || respuesta.status === 503) &&
                intento < 3
            ) {
                const retry = Number(
                    respuesta.headers.get("Retry-After")
                );

                const espera = Math.min(
                    (retry || 2 ** intento) * 1000,
                    10000
                );

                await new Promise(resolve => setTimeout(resolve, espera));

                return graphFetch(url, opciones, intento + 1);
            }

            if (respuesta.status === 401 && intento < 1) {
                await obtenerTokenSilencioso();
                return graphFetch(url, opciones, intento + 1);
            }

            return respuesta;
        } finally {
            clearTimeout(temporizador);
        }
    }

    async function graphJSON(url, opciones = {}) {
        const respuesta = await graphFetch(url, opciones);

        if (!respuesta.ok) {
            let detalle = "";

            try {
                const error = await respuesta.json();
                detalle = error?.error?.message || "";
            } catch (_) {}

            if (respuesta.status === 403) {
                throw new Error(
                    "Microsoft Graph denegó el acceso. Revisa los permisos y el consentimiento."
                );
            }

            if (respuesta.status === 404) {
                throw new Error(
                    "No se encontró la lista, carpeta o recurso de SharePoint."
                );
            }

            throw new Error(
                `Error de Microsoft Graph (${respuesta.status}). ` +
                (detalle || "Comprueba los permisos y la configuración.")
            );
        }

        if (respuesta.status === 204) {
            return null;
        }

        return respuesta.json();
    }

    async function obtenerTodosLosElementos(url) {
        const elementos = [];

        while (url) {
            const datos = await graphJSON(url);
            elementos.push(...(datos.value || []));
            url = datos["@odata.nextLink"] || null;
        }

        return elementos;
    }

    async function obtenerSitio(host, ruta) {
        const url = `${GRAPH}/sites/${host}:${ruta}?$select=id,webUrl`;
        return graphJSON(url);
    }

    /* =====================================================
       LISTA DE CONVERSACIONES
    ===================================================== */

    async function obtenerListaConversaciones() {
        if (listaConversaciones) {
            return listaConversaciones;
        }

        if (!sitioLista) {
            sitioLista = await obtenerSitio(
                LIST_SITE_HOST,
                LIST_SITE_PATH
            );
        }

        const url =
            `${GRAPH}/sites/${sitioLista.id}/lists` +
            `?$select=id,displayName,list&$top=200`;

        const listas = await obtenerTodosLosElementos(url);

        listaConversaciones = listas.find(
            lista =>
                normalizarTexto(lista.displayName) ===
                normalizarTexto(LIST_NAME)
        );

        if (!listaConversaciones) {
            throw new Error(
                `No se encontró la lista ${LIST_NAME} en el sitio configurado.`
            );
        }

        return listaConversaciones;
    }

    async function obtenerColumnasLista() {
        if (columnasLista) {
            return columnasLista;
        }

        const lista = await obtenerListaConversaciones();

        const url =
            `${GRAPH}/sites/${sitioLista.id}/lists/${lista.id}/columns` +
            "?$select=id,name,displayName";

        const columnas = await obtenerTodosLosElementos(url);

        columnasLista = {};

        columnas.forEach(columna => {
            columnasLista[
                normalizarNombreColumna(columna.displayName)
            ] = columna.name;

            columnasLista[
                normalizarNombreColumna(columna.name)
            ] = columna.name;
        });

        return columnasLista;
    }

    function nombreInternoColumna(nombre) {
        return columnasLista?.[normalizarNombreColumna(nombre)] || null;
    }

    async function guardarConversacionEnSharePoint(
        mensajeUsuario,
        respuestaBot,
        documentoOrigen,
        tema
    ) {
        const lista = await obtenerListaConversaciones();
        await obtenerColumnasLista();

        const requeridas = [
            "ConversationId",
            "UsuarioCorreo",
            "MensajeUsuario",
            "RespuestaBot",
            "Fecha",
            "DocumentoOrigen",
            "Tema"
        ];

        const campos = {};

        for (const nombre of requeridas) {
            const interno = nombreInternoColumna(nombre);

            if (!interno) {
                throw new Error(
                    `No se encontró la columna "${nombre}" en ${LIST_NAME}.`
                );
            }
        }

        const titleColumna =
            nombreInternoColumna("Title") ||
            nombreInternoColumna("Título");

        if (!titleColumna) {
            throw new Error(
                "No se encontró la columna predeterminada Title/Título."
            );
        }

        campos[titleColumna] =
            `Chat ${new Date().toLocaleString("es-PE")}`;

        campos[nombreInternoColumna("ConversationId")] = conversationId;
        campos[nombreInternoColumna("UsuarioCorreo")] = usuarioCorreo;
        campos[nombreInternoColumna("MensajeUsuario")] = mensajeUsuario;
        campos[nombreInternoColumna("RespuestaBot")] = respuestaBot;
        campos[nombreInternoColumna("Fecha")] = new Date().toISOString();
        campos[nombreInternoColumna("DocumentoOrigen")] =
            documentoOrigen || "";
        campos[nombreInternoColumna("Tema")] = tema || "Procedimientos TI";

        const url =
            `${GRAPH}/sites/${sitioLista.id}/lists/${lista.id}/items`;

        await graphJSON(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({ fields: campos })
        });
    }

    async function cargarHistorialDesdeSharePoint() {
        const lista = await obtenerListaConversaciones();
        await obtenerColumnasLista();

        const campoConversation =
            nombreInternoColumna("ConversationId");
        const campoUsuario = nombreInternoColumna("UsuarioCorreo");

        if (!campoConversation || !campoUsuario) {
            throw new Error("Faltan columnas para consultar el historial.");
        }

        // Filtrado en Graph para solicitar únicamente la conversación
        // del usuario autenticado.
        const filtro =
            `fields/${campoConversation} eq '${escaparOData(conversationId)}'` +
            ` and fields/${campoUsuario} eq '${escaparOData(usuarioCorreo)}'`;

        const url =
            `${GRAPH}/sites/${sitioLista.id}/lists/${lista.id}/items` +
            `?$expand=fields&$filter=${encodeURIComponent(filtro)}` +
            "&$top=100";

        const datos = await obtenerTodosLosElementos(url);

        const campoMensaje = nombreInternoColumna("MensajeUsuario");
        const campoRespuesta = nombreInternoColumna("RespuestaBot");
        const campoFecha = nombreInternoColumna("Fecha");

        const mensajes = [];

        datos.forEach(item => {
            const fields = item.fields || {};

            if (fields[campoMensaje]) {
                mensajes.push({
                    role: "user",
                    content: String(fields[campoMensaje]),
                    timestamp: Date.parse(fields[campoFecha] || "") || 0
                });
            }

            if (fields[campoRespuesta]) {
                mensajes.push({
                    role: "assistant",
                    content: String(fields[campoRespuesta]),
                    timestamp: (Date.parse(fields[campoFecha] || "") || 0) + 1
                });
            }
        });

        mensajes.sort((a, b) => a.timestamp - b.timestamp);

        historial = mensajes;
        guardarHistorial();

        return mensajes;
    }

    /* =====================================================
       SHAREPOINT: DOCUMENTOS PDF
    ===================================================== */

    async function obtenerDriveDocumentos() {
        if (driveDocumentos) {
            return driveDocumentos;
        }

        if (!sitioDocumentos) {
            sitioDocumentos = await obtenerSitio(
                DOCS_SITE_HOST,
                DOCS_SITE_PATH
            );
        }

        const url =
            `${GRAPH}/sites/${sitioDocumentos.id}/drives` +
            "?$select=id,name,driveType,webUrl";

        const drives = await obtenerTodosLosElementos(url);

        // La biblioteca puede mostrarse con el nombre localizado
        // o con su nombre interno en inglés.
        driveDocumentos = drives.find(drive =>
            /documentos compartidos|shared documents|documents/i.test(
                drive.name || ""
            )
        );

        if (!driveDocumentos && drives.length === 1) {
            driveDocumentos = drives[0];
        }

        if (!driveDocumentos) {
            throw new Error(
                "No se pudo identificar la biblioteca de documentos de TI."
            );
        }

        return driveDocumentos;
    }

    async function obtenerPDFsDeCarpeta() {
        const drive = await obtenerDriveDocumentos();

        const ruta =
            `${GRAPH}/drives/${drive.id}/root:/` +
            `${encodeURIComponent(DOCS_FOLDER)}` +
            ":/children?$select=id,name,size,file,webUrl,parentReference&$top=200";

        const elementos = await obtenerTodosLosElementos(ruta);

        archivosPDF = elementos.filter(item =>
            item.file &&
            /\.pdf$/i.test(item.name || "")
        );

        if (!archivosPDF.length) {
            throw new Error(
                `No se encontraron PDF directamente dentro de ${DOCS_FOLDER}.`
            );
        }

        return archivosPDF.slice(0, MAX_PDFS);
    }

    async function cargarPDFJS() {
        if (window.pdfjsLib) {
            return window.pdfjsLib;
        }

        await cargarScript(
            "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"
        );

        if (!window.pdfjsLib) {
            throw new Error("No se pudo inicializar PDF.js.");
        }

        window.pdfjsLib.GlobalWorkerOptions.workerSrc =
            "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

        return window.pdfjsLib;
    }

    async function extraerTextoPDF(archivo) {
        const drive = await obtenerDriveDocumentos();
        const url =
            `${GRAPH}/drives/${drive.id}/items/${archivo.id}/content`;

        const respuesta = await graphFetch(url, {
            timeoutMs: 90000
        });

        if (!respuesta.ok) {
            throw new Error(
                `No se pudo descargar el PDF "${archivo.name}".`
            );
        }

        const buffer = await respuesta.arrayBuffer();
        const pdfjs = await cargarPDFJS();

        const documento = await pdfjs.getDocument({
            data: buffer
        }).promise;

        const paginas = [];

        for (let numero = 1; numero <= documento.numPages; numero++) {
            const pagina = await documento.getPage(numero);
            const contenido = await pagina.getTextContent();

            const textoPagina = contenido.items
                .map(item => item.str || "")
                .join(" ");

            paginas.push(textoPagina);
        }

        return paginas.join("\n").replace(/\s+/g, " ").trim();
    }

    async function indexarDocumentos() {
        if (indiceDocumentos.length) {
            return indiceDocumentos;
        }

        if (promesaIndexacion) {
            return promesaIndexacion;
        }

        promesaIndexacion = (async () => {
            const archivos = await obtenerPDFsDeCarpeta();
            const documentos = [];

            for (const archivo of archivos) {
                try {
                    const texto = await extraerTextoPDF(archivo);

                    if (texto) {
                        documentos.push({
                            id: archivo.id,
                            nombre: archivo.name,
                            texto,
                            normalizado: normalizarTexto(texto),
                            url: archivo.webUrl || ""
                        });
                    }
                } catch (error) {
                    console.warn(
                        "[ALFERZA] No se pudo procesar un PDF:",
                        archivo.name,
                        error.message
                    );
                }
            }

            indiceDocumentos = documentos;

            if (!documentos.length) {
                throw new Error(
                    "No se pudo extraer texto de los PDF. " +
                    "Comprueba los permisos y si los archivos contienen texto seleccionable."
                );
            }

            return documentos;
        })();

        try {
            return await promesaIndexacion;
        } finally {
            promesaIndexacion = null;
        }
    }

    /* =====================================================
       BÚSQUEDA LOCAL EN LOS PDF
    ===================================================== */

    const PALABRAS_VACIAS = new Set([
        "como", "que", "para", "por", "con", "sin", "una", "uno",
        "unos", "unas", "los", "las", "del", "de", "el", "la",
        "y", "o", "en", "un", "al", "es", "son", "se", "me",
        "mi", "tu", "su", "lo", "le", "les", "necesito", "quiero",
        "puedo", "puede", "donde", "cuando", "cual", "cuales",
        "hay", "tiene", "tienen", "hacer", "favor", "hola",
        "buenos", "buenas", "dias", "tardes", "noches"
    ]);

    function extraerTerminos(pregunta) {
        return [...new Set(
            normalizarTexto(pregunta)
                .replace(/[^a-z0-9\s.-]/g, " ")
                .split(/\s+/)
                .filter(palabra =>
                    palabra.length >= 3 &&
                    !PALABRAS_VACIAS.has(palabra)
                )
        )];
    }

    function buscarEnDocumentos(pregunta, documentos) {
        const terminos = extraerTerminos(pregunta);

        if (!terminos.length) {
            return null;
        }

        const resultados = documentos.map(documento => {
            let puntuacion = 0;

            for (const termino of terminos) {
                let posicion = 0;
                let apariciones = 0;

                while (
                    (posicion = documento.normalizado.indexOf(
                        termino,
                        posicion
                    )) !== -1
                ) {
                    apariciones++;
                    posicion += termino.length;

                    if (apariciones >= 8) break;
                }

                puntuacion += Math.min(apariciones, 5);
            }

            const cobertura = terminos.filter(termino =>
                documento.normalizado.includes(termino)
            ).length;

            return {
                documento,
                puntuacion,
                cobertura,
                porcentaje: cobertura / terminos.length
            };
        }).filter(resultado =>
            resultado.puntuacion > 0 &&
            resultado.porcentaje >= 0.3
        ).sort((a, b) =>
            b.puntuacion - a.puntuacion ||
            b.porcentaje - a.porcentaje
        );

        if (!resultados.length) {
            return null;
        }

        const mejor = resultados[0];
        const texto = mejor.documento.texto;
        const normalizado = mejor.documento.normalizado;

        let mejorPosicion = -1;
        let mejorVentana = -1;

        for (const termino of terminos) {
            const posicion = normalizado.indexOf(termino);

            if (posicion < 0) continue;

            const inicio = Math.max(0, posicion - 250);
            const fin = Math.min(texto.length, posicion + 1100);
            const ventana = texto.slice(inicio, fin);

            const puntuacionVentana = terminos.filter(otro =>
                normalizarTexto(ventana).includes(otro)
            ).length;

            if (puntuacionVentana > mejorVentana) {
                mejorVentana = puntuacionVentana;
                mejorPosicion = inicio;
            }
        }

        if (mejorPosicion < 0) {
            mejorPosicion = 0;
        }

        let extracto = texto.slice(
            mejorPosicion,
            mejorPosicion + 1500
        ).trim();

        if (mejorPosicion > 0) {
            extracto = "… " + extracto;
        }

        if (mejorPosicion + 1500 < texto.length) {
            extracto += " …";
        }

        return {
            text:
                "Encontré información relacionada en un procedimiento de ALFERZA:\n\n" +
                extracto +
                "\n\nDocumento: " + mejor.documento.nombre +
                "\n\nEsta respuesta se basa en coincidencias de texto. " +
                "Verifica el procedimiento completo antes de ejecutar acciones críticas.",
            source: mejor.documento.nombre,
            score: mejor.puntuacion,
            mode: "documentos"
        };
    }

    /* =====================================================
       RESPUESTAS GENERALES SIN IA
    ===================================================== */

    function responderSaludo(pregunta) {
        const texto = normalizarTexto(pregunta);

        if (/^(hola|buenas|buenos dias|buenas tardes|buenas noches|hey|que tal)[!. ]*$/.test(texto)) {
            return {
                text:
                    "¡Hola! Soy el asistente de TI de ALFERZA. " +
                    "Puedo buscar información en los procedimientos PDF disponibles. " +
                    "¿Qué necesitas consultar?",
                mode: "general"
            };
        }

        if (
            texto.includes("que puedes hacer") ||
            texto.includes("como puedes ayudar")
        ) {
            return {
                text:
                    "Puedo buscar coincidencias en los procedimientos PDF de TI " +
                    "y guardar esta conversación en SharePoint. " +
                    "No tengo razonamiento de IA en esta versión; si no encuentro " +
                    "información suficiente, te lo indicaré.",
                mode: "general"
            };
        }

        return null;
    }

    /* =====================================================
       CACHÉ
    ===================================================== */

    function obtenerRespuestaCache(pregunta) {
        const clave = normalizarTexto(pregunta);
        const entrada = cache[clave];

        if (!entrada || typeof entrada.text !== "string") {
            return null;
        }

        return entrada;
    }

    function guardarRespuestaCache(pregunta, respuesta) {
        if (!respuesta?.text || respuesta.mode === "error") {
            return;
        }

        cache[normalizarTexto(pregunta)] = {
            text: respuesta.text,
            mode: respuesta.mode || "general",
            source: respuesta.source || "",
            savedAt: Date.now()
        };

        const claves = Object.keys(cache).sort(
            (a, b) => cache[a].savedAt - cache[b].savedAt
        );

        while (claves.length > MAX_ENTRADAS_CACHE) {
            delete cache[claves.shift()];
        }

        guardarCache();
    }

    /* =====================================================
       HISTORIAL LOCAL
    ===================================================== */

    function registrarHistorial(role, content) {
        historial.push({
            role,
            content: String(content || ""),
            timestamp: Date.now()
        });

        guardarHistorial();
    }

    /* =====================================================
       INTERFAZ
    ===================================================== */

    function instalarEstilos() {
        if (document.getElementById("alferza-chat-dynamic-styles")) {
            return;
        }

        const estilos = document.createElement("style");
        estilos.id = "alferza-chat-dynamic-styles";

        estilos.textContent = `
            #typingIndicator {
                transition: opacity .2s ease;
            }

            #typingIndicator:not([hidden]) {
                animation: alferzaTypingPulse 1.2s ease-in-out infinite;
            }

            .alferza-typing-dots {
                display: inline-flex;
                align-items: center;
                gap: 5px;
                margin-left: 5px;
            }

            .alferza-typing-dots span {
                width: 6px;
                height: 6px;
                border-radius: 50%;
                background: currentColor;
                animation: alferzaDotBounce 1s infinite ease-in-out;
            }

            .alferza-typing-dots span:nth-child(2) {
                animation-delay: .15s;
            }

            .alferza-typing-dots span:nth-child(3) {
                animation-delay: .3s;
            }

            .message-bubble,
            .message-bubble.limited-response {
                display: block;
                height: auto;
                max-height: none;
                overflow: visible;
                -webkit-line-clamp: unset;
                white-space: pre-wrap;
                overflow-wrap: anywhere;
                word-break: normal;
            }

            .message-row {
                max-width: 98%;
            }

            @keyframes alferzaTypingPulse {
                0%, 100% { opacity: .65; }
                50% { opacity: 1; }
            }

            @keyframes alferzaDotBounce {
                0%, 60%, 100% {
                    transform: translateY(0);
                    opacity: .45;
                }
                30% {
                    transform: translateY(-4px);
                    opacity: 1;
                }
            }

            @media (prefers-reduced-motion: reduce) {
                #typingIndicator,
                .alferza-typing-dots span {
                    animation: none !important;
                    transition: none !important;
                }
            }
        `;

        document.head.appendChild(estilos);

        if (!elements.typing.querySelector(".alferza-typing-dots")) {
            const puntos = document.createElement("span");
            puntos.className = "alferza-typing-dots";
            puntos.setAttribute("aria-hidden", "true");

            for (let i = 0; i < 3; i++) {
                puntos.appendChild(document.createElement("span"));
            }

            elements.typing.appendChild(puntos);
        }

        elements.typing.setAttribute("role", "status");
        elements.typing.setAttribute(
            "aria-label",
            "ALFERZA está buscando información"
        );
    }

    function desplazarAlFinal() {
        requestAnimationFrame(() => {
            elements.messages.scrollTop =
                elements.messages.scrollHeight;
        });
    }

    function obtenerHora() {
        return new Intl.DateTimeFormat("es-PE", {
            hour: "2-digit",
            minute: "2-digit",
            hour12: false
        }).format(new Date());
    }

    function agregarMensaje(tipo, texto) {
        const fila = document.createElement("div");
        fila.className = `message-row ${tipo}`;

        const contenido = document.createElement("div");
        contenido.className = "message-content";

        const nombre = document.createElement("span");
        nombre.className = "message-name";
        nombre.textContent = tipo === "user"
            ? "Tú"
            : "Asistente ALFERZA";

        const burbuja = document.createElement("div");
        burbuja.className = "message-bubble";
        burbuja.textContent = String(texto || "");

        const hora = document.createElement("span");
        hora.className = "message-time";
        hora.textContent = obtenerHora();

        contenido.append(nombre, burbuja, hora);

        if (tipo === "bot") {
            const avatar = document.createElement("div");
            avatar.className = "message-avatar";
            avatar.textContent = "A";
            avatar.setAttribute("aria-hidden", "true");
            fila.appendChild(avatar);
        }

        fila.appendChild(contenido);
        elements.messages.appendChild(fila);

        desplazarAlFinal();
        return fila;
    }

    function mostrarEscritura(visible) {
        elements.typing.hidden = !visible;
        elements.sendButton.disabled = visible;
        elements.input.disabled = visible;

        elements.suggestions.forEach(boton => {
            boton.disabled = visible;
        });

        if (visible) {
            desplazarAlFinal();
        }
    }

    function actualizarContador() {
        elements.characterCount.textContent =
            `${elements.input.value.length}/${MAX_CARACTERES}`;

        elements.input.style.height = "auto";
        elements.input.style.height =
            `${Math.min(elements.input.scrollHeight, 120)}px`;
    }

    /* =====================================================
       PROCESAMIENTO DE CONSULTAS
    ===================================================== */

    async function consultarDocumentos(mensaje) {
        const respuestaGeneral = responderSaludo(mensaje);

        if (respuestaGeneral) {
            return respuestaGeneral;
        }

        const documentos = await indexarDocumentos();
        const resultado = buscarEnDocumentos(mensaje, documentos);

        if (resultado) {
            return resultado;
        }

        return {
            text:
                "No encontré coincidencias suficientes en los PDF de " +
                DOCS_FOLDER + " para responder con seguridad.\n\n" +
                "Prueba con el nombre del procedimiento, del sistema o del equipo. " +
                "También puedes consultar directamente el documento en SharePoint.",
            mode: "sin_coincidencias"
        };
    }

    async function enviarMensaje(mensaje) {
        const texto = String(mensaje || "").trim();

        if (!texto || procesandoMensaje) {
            return;
        }

        if (texto.length > MAX_CARACTERES) {
            agregarMensaje(
                "bot",
                `La consulta supera el límite de ${MAX_CARACTERES} caracteres.`
            );
            return;
        }

        procesandoMensaje = true;
        agregarMensaje("user", texto);

        registrarHistorial("user", texto);

        elements.input.value = "";
        actualizarContador();
        mostrarEscritura(true);

        let respuesta;
        let documentoOrigen = "";

        try {
            await obtenerTokenSilencioso();

            const cacheada = obtenerRespuestaCache(texto);

            if (cacheada) {
                respuesta = {
                    text: cacheada.text,
                    mode: cacheada.mode,
                    source: cacheada.source || ""
                };
            } else {
                respuesta = await consultarDocumentos(texto);
            }

            documentoOrigen = respuesta.source || "";

            // Guardar la pregunta y la respuesta en Microsoft Lists.
            await guardarConversacionEnSharePoint(
                texto,
                respuesta.text,
                documentoOrigen,
                respuesta.mode
            );

            agregarMensaje("bot", respuesta.text);
            registrarHistorial("assistant", respuesta.text);
            guardarRespuestaCache(texto, respuesta);

        } catch (error) {
            console.error("[ALFERZA CHATBOT]", error);

            let mensajeError =
                "No pude completar la consulta.\n\n" +
                "Detalle: " + (error.message || "Error desconocido") +
                "\n\nComprueba tu sesión Microsoft y los permisos de SharePoint.";

            if (
                /interaction_required|consent_required|login_required/i.test(
                    String(error.name) + " " + String(error.message)
                )
            ) {
                mensajeError =
                    "Necesitas autorizar el acceso a Microsoft Graph. " +
                    "Pulsa el botón de conexión Microsoft que aparecerá abajo.";
                mostrarBotonConexion();
            }

            agregarMensaje("bot", mensajeError);
        } finally {
            mostrarEscritura(false);
            procesandoMensaje = false;

            elements.input.disabled = false;
            elements.sendButton.disabled = false;

            actualizarContador();
            desplazarAlFinal();
            elements.input.focus();
        }
    }

    /* =====================================================
       CONEXIÓN INTERACTIVA
    ===================================================== */

    function mostrarBotonConexion() {
        if (document.getElementById("alferzaConnectMicrosoft")) {
            return;
        }

        const boton = document.createElement("button");
        boton.id = "alferzaConnectMicrosoft";
        boton.type = "button";
        boton.className = "suggestion-button";
        boton.textContent = "Conectar con Microsoft";

        boton.addEventListener("click", async () => {
            boton.disabled = true;

            try {
                await iniciarSesionInteractiva();
                boton.remove();

                agregarMensaje(
                    "bot",
                    "Conexión Microsoft realizada. Ya puedes volver a enviar tu consulta."
                );
            } catch (error) {
                console.error("[ALFERZA MSAL]", error);
                agregarMensaje(
                    "bot",
                    "No se pudo completar la conexión Microsoft. " +
                    "Revisa el registro de la aplicación y los permisos de Graph."
                );
            } finally {
                if (boton.isConnected) {
                    boton.disabled = false;
                }
            }
        });

        elements.messages.appendChild(boton);
        desplazarAlFinal();
    }

    /* =====================================================
       NUEVA CONVERSACIÓN
    ===================================================== */

    elements.clearButton.addEventListener("click", function () {
        if (procesandoMensaje) {
            return;
        }

        conversationId = generarId();
        historial = [];
        cache = {};

        try {
            sessionStorage.setItem(STORAGE_KEY, conversationId);
        } catch (_) {}

        guardarHistorial();
        guardarCache();

        elements.messages.replaceChildren();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "Has iniciado una conversación nueva. " +
            "Puedo buscar coincidencias en los procedimientos PDF disponibles."
        );

        elements.input.value = "";
        actualizarContador();
        elements.input.focus();
    });

    /* =====================================================
       EVENTOS
    ===================================================== */

    elements.form.addEventListener("submit", function (evento) {
        evento.preventDefault();
        enviarMensaje(elements.input.value);
    });

    elements.input.addEventListener("input", actualizarContador);

    elements.input.addEventListener("keydown", function (evento) {
        if (
            evento.key === "Enter" &&
            !evento.shiftKey &&
            !evento.isComposing
        ) {
            evento.preventDefault();

            if (!procesandoMensaje) {
                elements.form.requestSubmit();
            }
        }
    });

    elements.suggestions.forEach(boton => {
        boton.addEventListener("click", function () {
            const pregunta = boton.dataset.question;

            if (pregunta) {
                enviarMensaje(pregunta);
            }
        });
    });

    /* =====================================================
       INICIALIZACIÓN
    ===================================================== */

    async function iniciarChatbot() {
        instalarEstilos();
        actualizarContador();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "Estoy preparado para buscar información en los procedimientos PDF. " +
            "La primera consulta puede tardar mientras se cargan los documentos."
        );

        try {
            await inicializarMSAL();

            if (!cuentaActual) {
                agregarMensaje(
                    "bot",
                    "No encuentro una cuenta Microsoft disponible para Graph. " +
                    "Si ya iniciaste sesión en ALFERZA y sigue apareciendo este mensaje, " +
                    "conecta tu cuenta con el botón de abajo."
                );

                mostrarBotonConexion();
                return;
            }

            await obtenerTokenSilencioso();

            // Recuperar el historial guardado para esta conversación.
            // Si la consulta filtrada no está soportada por la lista,
            // se conserva el historial local y se registra el error.
            try {
                const mensajes = await cargarHistorialDesdeSharePoint();

                if (mensajes.length) {
                    elements.messages.replaceChildren();

                    agregarMensaje(
                        "bot",
                        "He recuperado el historial de esta conversación."
                    );

                    mensajes.forEach(mensaje => {
                        agregarMensaje(
                            mensaje.role === "user" ? "user" : "bot",
                            mensaje.content
                        );
                    });
                }
            } catch (error) {
                console.warn(
                    "[ALFERZA] No se pudo recuperar el historial remoto:",
                    error.message
                );
            }

            console.info(
                "[ALFERZA CHATBOT] Listo.",
                {
                    correo: usuarioCorreo,
                    conversationId,
                    lista: LIST_NAME,
                    carpeta: DOCS_FOLDER
                }
            );

        } catch (error) {
            console.error("[ALFERZA CHATBOT] Inicio:", error);

            agregarMensaje(
                "bot",
                "No se pudo preparar la conexión con Microsoft Graph. " +
                "Verifica que la sesión Microsoft esté activa y que la aplicación " +
                "tenga permisos para consultar SharePoint."
            );

            mostrarBotonConexion();
        }

        desplazarAlFinal();
    }

    iniciarChatbot();
})();
