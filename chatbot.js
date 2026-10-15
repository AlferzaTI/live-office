
"use strict";

/* =========================================================
   ALFERZA LIVE OFFICE - ASISTENTE DE SOPORTE TI
   Sin backend Node.js y sin OpenAI.
   Microsoft Graph + SharePoint PDF + Microsoft Lists.
========================================================= */

(function () {
    const CONFIG = {
        clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
        tenantId: "dbab984f-4bb1-4b60-9dff-da59f54acdf1",

        redirectUri:
            window.location.hostname === "localhost"
                ? "http://localhost:5500/blank.html"
                : "https://alferzati.github.io/live-office/blank.html",

        scopes: [
            "User.Read",
            "Sites.Read.All",
            "Sites.ReadWrite.All"
        ],

        // Sitio que contiene los procedimientos PDF.
        documentos: {
            hostname: "alferzaholding.sharepoint.com",
            rutaSitio: "/sites/TI",
            carpeta: "Procedimientos-TI"
        },

        // Sitio personal donde se encuentra la lista.
        lista: {
            hostname: "alferzaholding-my.sharepoint.com",
            rutaSitio: "/personal/soporte1_alferza_pe",
            nombre: "ConversacionesCHATBOT"
        },

        pdfJsUrl:
            "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",

        pdfWorkerUrl:
            "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js",

        maxCaracteres: 1000,
        maxMensajesContexto: 6,
        maxResultados: 4,
        maxPaginasPdf: 200,
        cacheMs: 15 * 60 * 1000
    };

    const GRAPH = "https://graph.microsoft.com/v1.0";

    const STORAGE_KEY = "alferza_chatbot_conversation_id";
    const HISTORY_PREFIX = "alferza_chat_history_";
    const CACHE_PREFIX = "alferza_chat_cache_";

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
        console.error(
            "[ALFERZA CHATBOT] Faltan elementos HTML del chat."
        );
        return;
    }

    let msalApp = null;
    let cuenta = null;
    let tokenActual = null;
    let tokenExpira = 0;

    let usuario = null;
    let conversationId = obtenerConversationId();
    let historial = cargarHistorialLocal();
    let cache = cargarCache();

    let procesandoMensaje = false;
    let inicializado = false;

    let siteDocumentosId = null;
    let siteListaId = null;
    let driveDocumentosId = null;
    let listaId = null;

    let columnasLista = {};
    let documentosIndexados = [];
    let indexacionEnCurso = null;
    let pdfJsCargado = false;
    let ultimaCargaDocumentos = 0;

    /* =====================================================
       1. UTILIDADES Y ALMACENAMIENTO LOCAL
    ===================================================== */

    function leerStorage(clave, predeterminado) {
        try {
            const valor = sessionStorage.getItem(clave);
            return valor ? JSON.parse(valor) : predeterminado;
        } catch (_) {
            return predeterminado;
        }
    }

    function guardarStorage(clave, valor) {
        try {
            sessionStorage.setItem(clave, JSON.stringify(valor));
            return true;
        } catch (_) {
            return false;
        }
    }

    function generarId() {
        if (window.crypto && crypto.randomUUID) {
            return crypto.randomUUID();
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

    function cargarHistorialLocal() {
        const datos = leerStorage(
            HISTORY_PREFIX + conversationId,
            []
        );

        return Array.isArray(datos) ? datos : [];
    }

    function guardarHistorialLocal() {
        guardarStorage(claveHistorial(), historial);
    }

    function cargarCache() {
        const datos = leerStorage(CACHE_PREFIX + conversationId, {});

        return datos &&
            typeof datos === "object" &&
            !Array.isArray(datos)
                ? datos
                : {};
    }

    function guardarCache() {
        guardarStorage(claveCache(), cache);
    }

    function normalizar(texto) {
        return String(texto || "")
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[^a-z0-9\s._/-]/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    function escaparOData(valor) {
        return String(valor || "").replace(/'/g, "''");
    }

    function nombreInterno(nombre) {
        return String(nombre || "")
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[^a-z0-9]/g, "");
    }

    function esErrorPermisos(error) {
        const texto = String(error && error.message || "").toLowerCase();

        return texto.includes("consent") ||
            texto.includes("permission") ||
            texto.includes("forbidden") ||
            texto.includes("403") ||
            texto.includes("401") ||
            texto.includes("interaction_required");
    }

    function mensajeErrorAmigable(error) {
        const texto = String(error && error.message || "");

        if (texto.includes("interaction_required")) {
            return "Microsoft necesita que vuelvas a iniciar sesión. " +
                "Recarga la página e inténtalo nuevamente.";
        }

        if (texto.includes("403") || texto.toLowerCase().includes("forbidden")) {
            return "Microsoft ha denegado el acceso. Revisa los permisos " +
                "Sites.Read.All y Sites.ReadWrite.All de la aplicación y " +
                "el acceso del usuario a SharePoint.";
        }

        if (texto.includes("401")) {
            return "La sesión de Microsoft ha caducado. " +
                "Vuelve a iniciar sesión.";
        }

        return texto || "Se produjo un error desconocido.";
    }

    /* =====================================================
       2. MSAL Y MICROSOFT GRAPH
    ===================================================== */

    async function inicializarMSAL() {
        if (typeof msal === "undefined") {
            throw new Error(
                "No se ha cargado MSAL. Incluye msal-browser.min.js " +
                "antes de chatbot.js en tu HTML."
            );
        }

        msalApp = new msal.PublicClientApplication({
            auth: {
                clientId: CONFIG.clientId,
                authority:
                    "https://login.microsoftonline.com/" +
                    CONFIG.tenantId,
                redirectUri: CONFIG.redirectUri
            },
            cache: {
                cacheLocation: "sessionStorage",
                storeAuthStateInCookie: false
            }
        });

        if (typeof msalApp.initialize === "function") {
            await msalApp.initialize();
        }

        if (typeof msalApp.handleRedirectPromise === "function") {
            await msalApp.handleRedirectPromise();
        }

        const cuentas = msalApp.getAllAccounts();

        if (cuentas.length > 0) {
            cuenta = cuentas[0];
            msalApp.setActiveAccount(cuenta);
        } else {
            // Reutiliza la sesión existente de ALFERZA.
            // alferza_login no sustituye la autenticación de Graph.
            if (sessionStorage.getItem("alferza_login") !== "true") {
                throw new Error(
                    "No hay una sesión activa de ALFERZA. " +
                    "Inicia sesión en el portal."
                );
            }

            const respuesta = await msalApp.loginPopup({
                scopes: CONFIG.scopes
            });

            cuenta = respuesta.account;
            msalApp.setActiveAccount(cuenta);
        }

        await obtenerToken(true);

        const perfil = await graphFetch("/me?$select=id,displayName,mail,userPrincipalName");

        usuario = {
            id: perfil.id,
            nombre: perfil.displayName || "Usuario ALFERZA",
            correo: perfil.mail || perfil.userPrincipalName || ""
        };

        if (!usuario.correo) {
            throw new Error(
                "Microsoft no devolvió el correo del usuario. " +
                "No se puede guardar el historial de forma segura."
            );
        }

        console.info("[ALFERZA] Sesión Microsoft lista:", usuario.correo);
    }

    async function obtenerToken(forzar) {
        if (
            !forzar &&
            tokenActual &&
            Date.now() < tokenExpira - 60000
        ) {
            return tokenActual;
        }

        const cuentas = msalApp.getAllAccounts();

        if (!cuenta && cuentas.length) {
            cuenta = cuentas[0];
        }

        if (!cuenta) {
            throw new Error("No hay una cuenta Microsoft autenticada.");
        }

        try {
            const respuesta = await msalApp.acquireTokenSilent({
                scopes: CONFIG.scopes,
                account: cuenta
            });

            tokenActual = respuesta.accessToken;
            tokenExpira = respuesta.expiresOn
                ? respuesta.expiresOn.getTime()
                : Date.now() + 3000000;

            return tokenActual;
        } catch (error) {
            if (
                error instanceof msal.InteractionRequiredAuthError ||
                error.errorCode === "interaction_required"
            ) {
                const respuesta = await msalApp.acquireTokenPopup({
                    scopes: CONFIG.scopes,
                    account: cuenta
                });

                tokenActual = respuesta.accessToken;
                tokenExpira = respuesta.expiresOn
                    ? respuesta.expiresOn.getTime()
                    : Date.now() + 3000000;

                return tokenActual;
            }

            throw error;
        }
    }

    async function graphFetch(ruta, opciones, reintento) {
        opciones = opciones || {};
        reintento = reintento || false;

        const token = await obtenerToken(false);
        const cabeceras = new Headers(opciones.headers || {});

        cabeceras.set("Authorization", "Bearer " + token);

        if (
            opciones.body &&
            !(opciones.body instanceof Blob) &&
            !cabeceras.has("Content-Type")
        ) {
            cabeceras.set("Content-Type", "application/json");
        }

        const url = ruta.startsWith("https://")
            ? ruta
            : GRAPH + ruta;

        const respuesta = await fetch(url, {
            ...opciones,
            headers: cabeceras
        });

        if (respuesta.status === 401 && !reintento) {
            tokenActual = null;
            await obtenerToken(true);
            return graphFetch(ruta, opciones, true);
        }

        if (!respuesta.ok) {
            let detalle = "";

            try {
                const json = await respuesta.json();
                detalle = json.error && json.error.message
                    ? json.error.message
                    : JSON.stringify(json);
            } catch (_) {
                detalle = await respuesta.text().catch(() => "");
            }

            throw new Error(
                "Microsoft Graph HTTP " + respuesta.status +
                (detalle ? ": " + detalle : "")
            );
        }

        if (respuesta.status === 204) {
            return null;
        }

        const tipo = respuesta.headers.get("content-type") || "";

        if (tipo.includes("application/json")) {
            return respuesta.json();
        }

        return respuesta;
    }

    async function obtenerTodasLasPaginas(url) {
        const resultados = [];
        let siguiente = url;

        while (siguiente) {
            const pagina = await graphFetch(siguiente);
            resultados.push(...(pagina.value || []));
            siguiente = pagina["@odata.nextLink"] || null;
        }

        return resultados;
    }

    /* =====================================================
       3. RESOLVER SITIOS, LISTA Y COLUMNAS
    ===================================================== */

    async function resolverSitio(hostname, ruta) {
        const url =
            "/sites/" + hostname + ":" + ruta +
            "?$select=id,displayName,webUrl";

        const sitio = await graphFetch(url);
        return sitio.id;
    }

    async function prepararSharePoint() {
        siteDocumentosId = await resolverSitio(
            CONFIG.documentos.hostname,
            CONFIG.documentos.rutaSitio
        );

        siteListaId = await resolverSitio(
            CONFIG.lista.hostname,
            CONFIG.lista.rutaSitio
        );

        const drives = await obtenerTodasLasPaginas(
            "/sites/" + siteDocumentosId +
            "/drives?$select=id,name,webUrl,driveType"
        );

        let drive = drives.find(function (d) {
            return /documentos compartidos|shared documents/i.test(
                d.name || ""
            );
        });

        if (!drive) {
            const predeterminado = await graphFetch(
                "/sites/" + siteDocumentosId +
                "/drive?$select=id,name,webUrl"
            );

            drive = predeterminado;
        }

        if (!drive || !drive.id) {
            throw new Error(
                "No encontré la biblioteca Documentos compartidos " +
                "en el sitio de SharePoint /sites/TI."
            );
        }

        driveDocumentosId = drive.id;

        const listas = await obtenerTodasLasPaginas(
            "/sites/" + siteListaId +
            "/lists?$select=id,displayName,list"
        );

        const lista = listas.find(function (l) {
            return normalizar(l.displayName) ===
                normalizar(CONFIG.lista.nombre);
        });

        if (!lista) {
            throw new Error(
                "No se encontró la lista " + CONFIG.lista.nombre +
                ". Comprueba el nombre y los permisos de SharePoint."
            );
        }

        listaId = lista.id;

        const columnas = await obtenerTodasLasPaginas(
            "/sites/" + siteListaId +
            "/lists/" + listaId +
            "/columns?$select=id,name,displayName,hidden"
        );

        columnasLista = {};

        columnas.forEach(function (columna) {
            const clave = nombreInterno(columna.displayName);
            columnasLista[clave] = columna.name;
        });

        console.info("[ALFERZA] SharePoint preparado.", {
            documentosSite: siteDocumentosId,
            listaSite: siteListaId,
            lista: CONFIG.lista.nombre,
            columnas: Object.keys(columnasLista)
        });
    }

    function columna(nombre) {
        return columnasLista[nombreInterno(nombre)] || null;
    }

    /* =====================================================
       4. CARGAR PDF.JS
    ===================================================== */

    async function cargarPdfJs() {
        if (pdfJsCargado && window.pdfjsLib) {
            return;
        }

        if (!window.pdfjsLib) {
            await new Promise(function (resolve, reject) {
                const existente = document.querySelector(
                    'script[data-alferza-pdfjs="true"]'
                );

                if (existente) {
                    existente.addEventListener("load", resolve, { once: true });
                    existente.addEventListener("error", reject, { once: true });
                    return;
                }

                const script = document.createElement("script");
                script.src = CONFIG.pdfJsUrl;
                script.async = true;
                script.dataset.alferzaPdfjs = "true";
                script.onload = resolve;
                script.onerror = function () {
                    reject(new Error(
                        "No se pudo cargar PDF.js. Comprueba la conexión " +
                        "y que el CDN esté permitido."
                    ));
                };

                document.head.appendChild(script);
            });
        }

        if (!window.pdfjsLib) {
            throw new Error("PDF.js se cargó, pero no está disponible.");
        }

        window.pdfjsLib.GlobalWorkerOptions.workerSrc =
            CONFIG.pdfWorkerUrl;

        pdfJsCargado = true;
    }

    /* =====================================================
       5. BUSCAR PDF EN PROCEDIMIENTOS-TI
    ===================================================== */

    async function listarArchivosCarpeta() {
        const driveId = driveDocumentosId;
        const carpeta = CONFIG.documentos.carpeta;

        // Primero intentamos encontrar la carpeta en la raíz.
        let carpetaItem;

        try {
            carpetaItem = await graphFetch(
                "/drives/" + driveId +
                "/root:/" + encodeURIComponent(carpeta) +
                "?$select=id,name,folder,webUrl"
            );
        } catch (_) {
            const raiz = await obtenerTodasLasPaginas(
                "/drives/" + driveId +
                "/root/children?$select=id,name,folder,file,webUrl"
            );

            carpetaItem = raiz.find(function (item) {
                return item.folder &&
                    normalizar(item.name) === normalizar(carpeta);
            });
        }

        if (!carpetaItem || !carpetaItem.id) {
            throw new Error(
                "No encontré la carpeta " + carpeta +
                " en la biblioteca de documentos. Verifica la ruta."
            );
        }

        const archivos = [];
        const pendientes = [carpetaItem.id];
        const visitadas = new Set();

        while (pendientes.length) {
            const idCarpeta = pendientes.shift();

            if (visitadas.has(idCarpeta)) {
                continue;
            }

            visitadas.add(idCarpeta);

            const elementos = await obtenerTodasLasPaginas(
                "/drives/" + driveId +
                "/items/" + idCarpeta +
                "/children?$select=id,name,file,folder,webUrl,size"
            );

            for (const item of elementos) {
                if (item.folder) {
                    pendientes.push(item.id);
                } else if (
                    item.file &&
                    /\.pdf$/i.test(item.name || "")
                ) {
                    archivos.push(item);
                }
            }
        }

        return archivos;
    }

    async function extraerTextoPdf(archivo) {
        const respuesta = await graphFetch(
            "/drives/" + driveDocumentosId +
            "/items/" + archivo.id + "/content"
        );

        const buffer = await respuesta.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument({
            data: new Uint8Array(buffer)
        }).promise;

        const partes = [];
        const totalPaginas = Math.min(
            pdf.numPages,
            CONFIG.maxPaginasPdf
        );

        for (let numero = 1; numero <= totalPaginas; numero++) {
            const pagina = await pdf.getPage(numero);
            const contenido = await pagina.getTextContent();

            const texto = contenido.items
                .map(function (item) {
                    return item.str || "";
                })
                .join(" ")
                .replace(/\s+/g, " ")
                .trim();

            if (texto) {
                partes.push({
                    pagina: numero,
                    texto: texto
                });
            }
        }

        return partes;
    }

    function dividirEnFragmentos(texto, maximo) {
        const palabras = String(texto || "").split(/\s+/);
        const fragmentos = [];
        let actual = [];

        for (const palabra of palabras) {
            actual.push(palabra);

            if (actual.join(" ").length >= maximo) {
                fragmentos.push(actual.join(" "));
                actual = actual.slice(-35);
            }
        }

        if (actual.length) {
            fragmentos.push(actual.join(" "));
        }

        return fragmentos.filter(function (t) {
            return t.trim().length > 35;
        });
    }

    async function indexarDocumentos(forzar) {
        if (
            !forzar &&
            documentosIndexados.length &&
            Date.now() - ultimaCargaDocumentos < CONFIG.cacheMs
        ) {
            return documentosIndexados;
        }

        if (indexacionEnCurso) {
            return indexacionEnCurso;
        }

        indexacionEnCurso = (async function () {
            await cargarPdfJs();

            const archivos = await listarArchivosCarpeta();

            if (!archivos.length) {
                throw new Error(
                    "No hay archivos PDF en " +
                    CONFIG.documentos.carpeta + "."
                );
            }

            const indice = [];
            const errores = [];

            for (const archivo of archivos) {
                try {
                    const paginas = await extraerTextoPdf(archivo);

                    for (const pagina of paginas) {
                        const fragmentos = dividirEnFragmentos(
                            pagina.texto,
                            900
                        );

                        fragmentos.forEach(function (fragmento) {
                            indice.push({
                                archivo: archivo.name,
                                webUrl: archivo.webUrl || "",
                                pagina: pagina.pagina,
                                texto: fragmento,
                                normalizado: normalizar(fragmento)
                            });
                        });
                    }

                    console.info(
                        "[ALFERZA] PDF procesado:",
                        archivo.name,
                        "páginas:",
                        paginas.length
                    );
                } catch (error) {
                    errores.push(archivo.name);
                    console.warn(
                        "[ALFERZA] No se pudo leer el PDF:",
                        archivo.name,
                        error
                    );
                }
            }

            documentosIndexados = indice;
            ultimaCargaDocumentos = Date.now();

            if (!indice.length) {
                throw new Error(
                    "No se pudo extraer texto de los PDF. " +
                    "Si son documentos escaneados, necesitan OCR."
                );
            }

            console.info("[ALFERZA] Índice listo:", {
                archivos: archivos.length,
                fragmentos: indice.length,
                noLeidos: errores
            });

            return indice;
        })();

        try {
            return await indexacionEnCurso;
        } finally {
            indexacionEnCurso = null;
        }
    }

    /* =====================================================
       6. BÚSQUEDA LOCAL DE FRAGMENTOS RELEVANTES
    ===================================================== */

    const PALABRAS_IGNORADAS = new Set([
        "como", "que", "para", "por", "con", "sin", "una", "uno",
        "unos", "unas", "del", "las", "los", "este", "esta",
        "esto", "ese", "esa", "eso", "hay", "ser", "hacer",
        "tengo", "tiene", "puedo", "puede", "quiero", "necesito",
        "donde", "cuando", "cual", "cuales", "quien", "porque",
        "sobre", "entre", "desde", "hasta", "muy", "más", "mas",
        "mi", "mis", "tu", "tus", "su", "sus", "el", "la",
        "un", "y", "o", "a", "de", "en", "es", "al", "me",
        "te", "se", "lo", "le", "ya", "si", "no", "qué",
        "cómo", "por favor", "hola", "buenos", "buenas"
    ]);

    function palabrasClave(texto) {
        return [...new Set(
            normalizar(texto)
                .split(/\s+/)
                .filter(function (palabra) {
                    return palabra.length >= 2 &&
                        !PALABRAS_IGNORADAS.has(palabra);
                })
        )];
    }

    function buscarFragmentos(pregunta) {
        const terminos = palabrasClave(pregunta);

        if (!terminos.length) {
            return [];
        }

        const resultados = documentosIndexados.map(function (fragmento) {
            let puntuacion = 0;

            for (const termino of terminos) {
                if (fragmento.normalizado.includes(termino)) {
                    puntuacion += termino.length >= 7 ? 3 : 1;
                }
            }

            const cobertura = puntuacion / Math.max(terminos.length, 1);

            return {
                ...fragmento,
                puntuacion: puntuacion,
                cobertura: cobertura
            };
        }).filter(function (fragmento) {
            return fragmento.puntuacion > 0;
        });

        resultados.sort(function (a, b) {
            return b.puntuacion - a.puntuacion;
        });

        // Evitar devolver fragmentos casi idénticos del mismo PDF.
        const seleccionados = [];
        const paginasUsadas = new Set();

        for (const resultado of resultados) {
            const clave =
                resultado.archivo + "|" + resultado.pagina;

            if (paginasUsadas.has(clave)) {
                continue;
            }

            paginasUsadas.add(clave);
            seleccionados.push(resultado);

            if (seleccionados.length >= CONFIG.maxResultados) {
                break;
            }
        }

        return seleccionados;
    }

    function respuestaDesdeDocumentos(pregunta) {
        const resultados = buscarFragmentos(pregunta);

        if (!resultados.length || resultados[0].puntuacion < 2) {
            return {
                text:
                    "No encontré información suficientemente relacionada " +
                    "con tu consulta en los procedimientos PDF disponibles.\n\n" +
                    "Prueba con el nombre del sistema, equipo, error o " +
                    "procedimiento que necesitas. También puedes indicar " +
                    "más detalles del problema.",
                mode: "sin-resultados",
                fuentes: []
            };
        }

        const bloques = resultados.map(function (resultado, indice) {
            return (
                (indice + 1) + ". " + resultado.archivo +
                " (página " + resultado.pagina + ")\n" +
                resultado.texto
            );
        });

        const fuentes = resultados.map(function (resultado) {
            return {
                titulo: resultado.archivo,
                pagina: resultado.pagina,
                url: resultado.webUrl
            };
        });

        return {
            text:
                "Encontré información relacionada en tus procedimientos " +
                "de TI:\n\n" +
                bloques.join("\n\n") +
                "\n\nNota: estos son fragmentos del documento original. " +
                "Si necesitas pasos adicionales, dime qué parte quieres " +
                "aclarar.",
            mode: "sharepoint",
            fuentes: fuentes
        };
    }

    /* =====================================================
       7. HISTORIAL REMOTO EN MICROSOFT LISTS
    ===================================================== */

    function agregarHistorial(role, content) {
        historial.push({
            role: role,
            content: String(content || ""),
            timestamp: Date.now()
        });

        // El historial local sirve para el contexto inmediato.
        if (historial.length > 100) {
            historial = historial.slice(-100);
        }

        guardarHistorialLocal();
    }

    function historialContexto() {
        return historial
            .slice(-CONFIG.maxMensajesContexto)
            .map(function (mensaje) {
                return {
                    role: mensaje.role,
                    content: mensaje.content
                };
            });
    }

    function enriquecerPreguntaConContexto(pregunta) {
        const actual = normalizar(pregunta);

        // Si la pregunta parece un seguimiento corto, usar la última
        // pregunta del usuario para aportar contexto a la búsqueda.
        const seguimiento =
            /^(y |y si |y como |y cómo |eso |ese |esa |lo mismo|tambien|también|por que|por qué|donde|dónde|como hago|cómo hago|explica|amplia|amplía|detalla|continua|continúa)/i
                .test(actual) ||
            actual.split(/\s+/).length <= 3;

        if (!seguimiento) {
            return pregunta;
        }

        const ultimaPregunta = [...historial]
            .reverse()
            .find(function (mensaje) {
                return mensaje.role === "user";
            });

        if (!ultimaPregunta) {
            return pregunta;
        }

        return ultimaPregunta.content + " " + pregunta;
    }

    async function cargarHistorialRemoto() {
        const colConversation = columna("ConversationId");
        const colCorreo = columna("UsuarioCorreo");
        const colPregunta = columna("MensajeUsuario");
        const colRespuesta = columna("RespuestaBot");
        const colFecha = columna("Fecha");

        if (
            !colConversation ||
            !colCorreo ||
            !colPregunta ||
            !colRespuesta
        ) {
            console.warn(
                "[ALFERZA] Faltan columnas necesarias para recuperar " +
                "el historial remoto. Se conservará el historial local."
            );
            return;
        }

        // Filtrar en Graph por usuario y conversación.
        // No descargar conversaciones de otros usuarios.
        const filtro =
            "fields/" + colConversation + " eq '" +
            escaparOData(conversationId) + "' and fields/" +
            colCorreo + " eq '" + escaparOData(usuario.correo) + "'";

        const seleccion = [
            colConversation,
            colCorreo,
            colPregunta,
            colRespuesta,
            colFecha
        ].filter(Boolean).join(",");

        const url =
            "/sites/" + siteListaId +
            "/lists/" + listaId +
            "/items?$expand=fields($select=" +
            encodeURIComponent(seleccion) + ")" +
            "&$filter=" + encodeURIComponent(filtro) +
            "&$top=100";

        try {
            const registros = await obtenerTodasLasPaginas(url);
            const remoto = [];

            registros.forEach(function (registro) {
                const campos = registro.fields || {};
                const pregunta = campos[colPregunta];
                const respuesta = campos[colRespuesta];

                if (pregunta) {
                    remoto.push({
                        role: "user",
                        content: String(pregunta),
                        timestamp: Date.parse(campos[colFecha] || "") || 0
                    });
                }

                if (respuesta) {
                    remoto.push({
                        role: "assistant",
                        content: String(respuesta),
                        timestamp: Date.parse(campos[colFecha] || "") || 0
                    });
                }
            });

            remoto.sort(function (a, b) {
                return a.timestamp - b.timestamp;
            });

            if (remoto.length) {
                historial = remoto.slice(-100);
                guardarHistorialLocal();
            }
        } catch (error) {
            // No se consultan todos los registros como alternativa,
            // porque eso podría exponer el historial de otros usuarios.
            console.warn(
                "[ALFERZA] No se pudo recuperar el historial remoto " +
                "con el filtro por usuario y conversación:",
                error
            );
        }
    }

    async function guardarIntercambio(pregunta, respuesta, fuentes, tema) {
        const col = {
            titulo: columna("Title"),
            conversationId: columna("ConversationId"),
            correo: columna("UsuarioCorreo"),
            pregunta: columna("MensajeUsuario"),
            respuesta: columna("RespuestaBot"),
            fecha: columna("Fecha"),
            documento: columna("DocumentoOrigen"),
            tema: columna("Tema")
        };

        const campos = {};

        if (col.titulo) {
            campos[col.titulo] = pregunta.slice(0, 255);
        }

        if (col.conversationId) {
            campos[col.conversationId] = conversationId;
        }

        if (col.correo) {
            campos[col.correo] = usuario.correo;
        }

        if (col.pregunta) {
            campos[col.pregunta] = pregunta;
        }

        if (col.respuesta) {
            campos[col.respuesta] = respuesta;
        }

        if (col.fecha) {
            campos[col.fecha] = new Date().toISOString();
        }

        if (col.documento) {
            campos[col.documento] = (fuentes || [])
                .map(function (fuente) {
                    return fuente.titulo +
                        (fuente.pagina ? " (p. " + fuente.pagina + ")" : "");
                })
                .join("; ")
                .slice(0, 255);
        }

        if (col.tema) {
            campos[col.tema] = String(tema || "Soporte TI").slice(0, 255);
        }

        if (!col.pregunta || !col.respuesta || !col.correo || !col.conversationId) {
            throw new Error(
                "La lista ConversacionesCHATBOT no tiene todas las columnas " +
                "requeridas. Debe incluir UsuarioCorreo, ConversationId, " +
                "MensajeUsuario y RespuestaBot."
            );
        }

        await graphFetch(
            "/sites/" + siteListaId +
            "/lists/" + listaId + "/items",
            {
                method: "POST",
                body: JSON.stringify({
                    fields: campos
                })
            }
        );

        console.info("[ALFERZA] Intercambio guardado en Microsoft Lists.");
    }

    /* =====================================================
       8. CACHÉ DE RESPUESTAS
    ===================================================== */

    function obtenerRespuestaCache(pregunta) {
        const clave = normalizar(pregunta);
        const entrada = cache[clave];

        if (!entrada || typeof entrada.text !== "string") {
            return null;
        }

        if (
            !entrada.savedAt ||
            Date.now() - entrada.savedAt > CONFIG.cacheMs
        ) {
            delete cache[clave];
            guardarCache();
            return null;
        }

        return entrada;
    }

    function guardarRespuestaCache(pregunta, respuesta) {
        if (!respuesta || !respuesta.text) {
            return;
        }

        const clave = normalizar(pregunta);

        cache[clave] = {
            text: respuesta.text,
            mode: respuesta.mode || "sharepoint",
            fuentes: respuesta.fuentes || [],
            savedAt: Date.now()
        };

        const claves = Object.keys(cache);

        while (claves.length > 100) {
            delete cache[claves.shift()];
        }

        guardarCache();
    }

    /* =====================================================
       9. ESTILOS DINÁMICOS Y DESPLAZAMIENTO
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
                -webkit-box-orient: initial;
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
            "ALFERZA está buscando en los procedimientos"
        );
    }

    function desplazarAlFinal(inmediato) {
        const desplazar = function () {
            elements.messages.scrollTop =
                elements.messages.scrollHeight;

            if (
                !elements.typing.hidden &&
                elements.messages.contains(elements.typing)
            ) {
                elements.typing.scrollIntoView({
                    block: "nearest",
                    behavior: "auto"
                });
            }
        };

        if (inmediato) {
            desplazar();
        }

        requestAnimationFrame(function () {
            desplazar();
            requestAnimationFrame(desplazar);
        });
    }

    /* =====================================================
       10. MENSAJES DE LA INTERFAZ
    ===================================================== */

    function obtenerHora() {
        return new Intl.DateTimeFormat("es-PE", {
            hour: "2-digit",
            minute: "2-digit",
            hour12: false
        }).format(new Date());
    }

    function agregarMensaje(tipo, texto) {
        const fila = document.createElement("div");
        fila.className = "message-row " + tipo;

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
            fila.append(avatar);
        }

        fila.append(contenido);
        elements.messages.append(fila);

        desplazarAlFinal();

        return fila;
    }

    function mostrarEscritura(visible) {
        elements.typing.hidden = !visible;
        elements.sendButton.disabled = visible;
        elements.input.disabled = visible;

        elements.suggestions.forEach(function (boton) {
            boton.disabled = visible;
        });

        if (visible) {
            desplazarAlFinal();
        }
    }

    function actualizarContador() {
        const cantidad = elements.input.value.length;

        elements.characterCount.textContent =
            cantidad + "/" + CONFIG.maxCaracteres;

        elements.input.style.height = "auto";
        elements.input.style.height =
            Math.min(elements.input.scrollHeight, 120) + "px";
    }

    function mostrarFuentesComoEnlaces(fuentes) {
        if (!fuentes || !fuentes.length) {
            return "";
        }

        return "\n\nDocumentos consultados:\n" +
            fuentes.map(function (fuente) {
                const pagina = fuente.pagina
                    ? " (página " + fuente.pagina + ")"
                    : "";

                return "• " + fuente.titulo + pagina;
            }).join("\n");
    }

    /* =====================================================
       11. PROCESAR CONSULTA SIN BACKEND NI IA
    ===================================================== */

    async function consultarDocumentos(mensaje) {
        const preguntaConContexto =
            enriquecerPreguntaConContexto(mensaje);

        await indexarDocumentos(false);

        const resultado = respuestaDesdeDocumentos(
            preguntaConContexto
        );

        return {
            text: resultado.text +
                mostrarFuentesComoEnlaces(resultado.fuentes),
            mode: resultado.mode,
            fuentes: resultado.fuentes
        };
    }

    async function enviarMensaje(mensaje) {
        const texto = String(mensaje || "").trim();

        if (!texto || procesandoMensaje) {
            return;
        }

        if (texto.length > CONFIG.maxCaracteres) {
            agregarMensaje(
                "bot",
                "La consulta supera el límite de " +
                CONFIG.maxCaracteres + " caracteres."
            );
            return;
        }

        if (!inicializado) {
            agregarMensaje(
                "bot",
                "Estoy terminando de conectarme con Microsoft 365. " +
                "Espera unos segundos y vuelve a intentarlo."
            );
            return;
        }

        procesandoMensaje = true;

        agregarMensaje("user", texto);
        agregarHistorial("user", texto);

        elements.input.value = "";
        actualizarContador();
        mostrarEscritura(true);

        let respuestaFinal = null;

        try {
            const respuestaGuardada = obtenerRespuestaCache(texto);

            if (respuestaGuardada) {
                respuestaFinal = {
                    text: respuestaGuardada.text,
                    mode: respuestaGuardada.mode,
                    fuentes: respuestaGuardada.fuentes || []
                };
            } else {
                respuestaFinal = await consultarDocumentos(texto);
                guardarRespuestaCache(texto, respuestaFinal);
            }

            agregarMensaje("bot", respuestaFinal.text);
            agregarHistorial("assistant", respuestaFinal.text);

            // La respuesta se muestra aunque el guardado falle.
            try {
                await guardarIntercambio(
                    texto,
                    respuestaFinal.text,
                    respuestaFinal.fuentes,
                    respuestaFinal.mode
                );
            } catch (errorGuardar) {
                console.error(
                    "[ALFERZA] No se pudo guardar la conversación:",
                    errorGuardar
                );

                agregarMensaje(
                    "bot",
                    "Aviso: pude responder, pero no logré guardar " +
                    "este intercambio en Microsoft Lists. " +
                    mensajeErrorAmigable(errorGuardar)
                );
            }
        } catch (error) {
            console.error("[ALFERZA CHATBOT] Error:", error);

            let textoError =
                "No pude completar la consulta.\n\n" +
                mensajeErrorAmigable(error);

            if (
                String(error.message || "").includes("No hay archivos PDF") ||
                String(error.message || "").includes("No se pudo extraer texto")
            ) {
                textoError +=
                    "\n\nComprueba que los procedimientos estén en la carpeta " +
                    "Procedimientos-TI y que los PDF permitan seleccionar texto.";
            }

            agregarMensaje("bot", textoError);
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
       12. EVENTOS DEL CHAT
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

    elements.suggestions.forEach(function (boton) {
        boton.addEventListener("click", function () {
            const pregunta = boton.dataset.question;

            if (pregunta) {
                enviarMensaje(pregunta);
            }
        });
    });

    /* =====================================================
       13. NUEVA CONVERSACIÓN
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
        } catch (_) {
            // El chat sigue funcionando durante esta sesión.
        }

        guardarHistorialLocal();
        guardarCache();

        elements.messages.replaceChildren();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "He iniciado una conversación nueva. " +
            "Puedo buscar información en los procedimientos internos. " +
            "¿Qué necesitas resolver hoy?"
        );

        elements.input.value = "";
        actualizarContador();
        desplazarAlFinal();
        elements.input.focus();
    });

    /* =====================================================
       14. INICIALIZACIÓN
    ===================================================== */

    async function iniciarChatbot() {
        instalarEstilos();
        actualizarContador();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "Estoy conectando con Microsoft 365 para consultar " +
            "los procedimientos internos y guardar las conversaciones."
        );

        mostrarEscritura(true);

        try {
            await inicializarMSAL();
            await prepararSharePoint();
            await cargarHistorialRemoto();

            inicializado = true;

            agregarMensaje(
                "bot",
                "Conexión con Microsoft 365 establecida. " +
                "Ya puedes hacer consultas sobre los procedimientos " +
                "disponibles en SharePoint."
            );

            console.info("[ALFERZA CHATBOT] Listo.", {
                conversationId: conversationId,
                usuario: usuario.correo,
                mensajesGuardados: historial.length,
                respuestasEnCache: Object.keys(cache).length
            });
        } catch (error) {
            console.error(
                "[ALFERZA CHATBOT] Error de inicialización:",
                error
            );

            agregarMensaje(
                "bot",
                "No pude conectarme correctamente con Microsoft 365.\n\n" +
                mensajeErrorAmigable(error) +
                "\n\nRevisa la consola del navegador (F12 > Console) " +
                "para ver el detalle técnico."
            );
        } finally {
            mostrarEscritura(false);
            actualizarContador();
            desplazarAlFinal();
        }
    }

    iniciarChatbot();
})();
