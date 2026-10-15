"use strict";

/* =========================================================
   ALFERZA LIVE OFFICE - ASISTENTE DE SOPORTE TI
   HTML + CSS + JavaScript
   Microsoft Graph + SharePoint PDF + Microsoft Lists
   Sin backend Node.js. Sin OpenAI.
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

        documentos: {
            hostname: "alferzaholding.sharepoint.com",
            rutaSitio: "/sites/TI",
            carpeta: "Procedimientos-TI"
        },

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
        console.error("[ALFERZA] Faltan elementos HTML del chat.");
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
    let listaDisponible = false;

    let siteDocumentosId = null;
    let siteListaId = null;
    let driveDocumentosId = null;
    let listaId = null;
    let nombreSitioLista = "";
    let columnasLista = {};

    let documentosIndexados = [];
    let indexacionEnCurso = null;
    let pdfJsCargado = false;
    let ultimaCargaDocumentos = 0;

    /* =====================================================
       1. ALMACENAMIENTO Y UTILIDADES
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
        const datos = leerStorage(claveHistorial(), []);
        return Array.isArray(datos) ? datos : [];
    }

    function guardarHistorialLocal() {
        guardarStorage(claveHistorial(), historial);
    }

    function cargarCache() {
        const datos = leerStorage(claveCache(), {});
        return datos && typeof datos === "object" &&
            !Array.isArray(datos) ? datos : {};
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

    function nombreInterno(nombre) {
        return String(nombre || "")
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[^a-z0-9]/g, "");
    }

    function mensajeErrorAmigable(error) {
        const texto = String(error?.message || "");
        const minusculas = texto.toLowerCase();

        if (
            minusculas.includes("interaction_required") ||
            minusculas.includes("login_required")
        ) {
            return "Microsoft necesita que vuelvas a iniciar sesión. " +
                "Recarga la página e inténtalo nuevamente.";
        }

        if (
            texto.includes("403") ||
            minusculas.includes("forbidden")
        ) {
            return "Microsoft ha denegado el acceso. Comprueba los permisos " +
                "de Microsoft Graph y el acceso de tu usuario a SharePoint.";
        }

        if (texto.includes("401")) {
            return "La sesión de Microsoft ha caducado. Vuelve a iniciar sesión.";
        }

        if (
            minusculas.includes("failed to fetch") ||
            minusculas.includes("networkerror")
        ) {
            return "No se pudo conectar con Microsoft Graph. Comprueba " +
                "tu conexión y la configuración de la aplicación.";
        }

        return texto || "Se produjo un error desconocido.";
    }

    /* =====================================================
       2. AUTENTICACIÓN MSAL
    ===================================================== */

    async function inicializarMSAL() {
        if (typeof msal === "undefined") {
            throw new Error(
                "MSAL no está cargado. Incluye msal-browser.min.js " +
                "antes de chatbot.js en el HTML."
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

        if (cuentas.length) {
            cuenta =
                msalApp.getActiveAccount() ||
                cuentas[0];

            msalApp.setActiveAccount(cuenta);
        } else {
            if (sessionStorage.getItem("alferza_login") !== "true") {
                throw new Error(
                    "No hay una sesión activa de ALFERZA. Inicia sesión " +
                    "en el portal antes de abrir el asistente."
                );
            }

            const respuesta = await msalApp.loginPopup({
                scopes: CONFIG.scopes
            });

            cuenta = respuesta.account;
            msalApp.setActiveAccount(cuenta);
        }

        await obtenerToken(true);

        const perfil = await graphFetch(
            "/me?$select=id,displayName,mail,userPrincipalName"
        );

        usuario = {
            id: perfil.id,
            nombre: perfil.displayName || "Usuario ALFERZA",
            correo: perfil.mail || perfil.userPrincipalName || ""
        };

        if (!usuario.correo) {
            throw new Error(
                "Microsoft no devolvió el correo del usuario. " +
                "No se puede identificar al propietario del historial."
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

    /* =====================================================
       3. MICROSOFT GRAPH
    ===================================================== */

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
            // Leer el cuerpo una sola vez para no perder el detalle del error.
            const cuerpo = await respuesta.text().catch(() => "");
            let detalle = cuerpo;

            try {
                const json = JSON.parse(cuerpo);
                detalle = json.error?.message || cuerpo;
            } catch (_) {
                // La respuesta no era JSON.
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

    async function resolverSitio(hostname, ruta) {
        const rutaLimpia = ruta.startsWith("/")
            ? ruta
            : "/" + ruta;

        const url =
            "/sites/" + hostname + ":" + rutaLimpia +
            "?$select=id,displayName,webUrl";

        return graphFetch(url);
    }

    /* =====================================================
       4. RESOLVER SHAREPOINT Y BUSCAR LA LISTA
    ===================================================== */

    async function prepararSharePoint() {
        const sitioDocumentos = await resolverSitio(
            CONFIG.documentos.hostname,
            CONFIG.documentos.rutaSitio
        );

        siteDocumentosId = sitioDocumentos.id;

        console.info("[ALFERZA] Sitio de documentos:", {
            nombre: sitioDocumentos.displayName,
            url: sitioDocumentos.webUrl,
            id: sitioDocumentos.id
        });

        // Resolver la biblioteca de documentos.
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
            drive = await graphFetch(
                "/sites/" + siteDocumentosId +
                "/drive?$select=id,name,webUrl"
            );
        }

        if (!drive?.id) {
            throw new Error(
                "No encontré una biblioteca de documentos en " +
                sitioDocumentos.webUrl
            );
        }

        driveDocumentosId = drive.id;

        console.info("[ALFERZA] Biblioteca:", drive.name, drive.webUrl);

        // Buscar la lista primero en el sitio configurado y luego
        // en el sitio de documentos.
        const candidatos = [];
        const erroresSitios = [];

        try {
            const sitioLista = await resolverSitio(
                CONFIG.lista.hostname,
                CONFIG.lista.rutaSitio
            );

            candidatos.push(sitioLista);

            console.info("[ALFERZA] Sitio configurado para la lista:", {
                nombre: sitioLista.displayName,
                url: sitioLista.webUrl,
                id: sitioLista.id
            });
        } catch (error) {
            erroresSitios.push(
                "Sitio configurado: " + error.message
            );

            console.warn(
                "[ALFERZA] No se pudo resolver el sitio configurado para " +
                "la lista:",
                error
            );
        }

        // Evitar consultar dos veces el mismo sitio.
        if (
            siteDocumentosId &&
            !candidatos.some(s => s.id === siteDocumentosId)
        ) {
            candidatos.push(sitioDocumentos);
        }

        let encontrada = null;
        const diagnostico = [];

        for (const sitio of candidatos) {
            try {
                const listas = await obtenerTodasLasPaginas(
                    "/sites/" + sitio.id +
                    "/lists?$select=id,displayName,webUrl,list"
                );

                const nombres = listas.map(l => l.displayName || "(sin nombre)");

                diagnostico.push({
                    sitio: sitio.webUrl,
                    listas: nombres
                });

                console.info(
                    "[ALFERZA] Listas encontradas en",
                    sitio.webUrl,
                    nombres
                );

                encontrada = listas.find(function (lista) {
                    return normalizar(lista.displayName) ===
                        normalizar(CONFIG.lista.nombre);
                });

                if (encontrada) {
                    siteListaId = sitio.id;
                    nombreSitioLista = sitio.webUrl;
                    break;
                }
            } catch (error) {
                erroresSitios.push(
                    sitio.webUrl + ": " + error.message
                );

                console.error(
                    "[ALFERZA] Error al enumerar las listas de",
                    sitio.webUrl,
                    error
                );
            }
        }

        if (!encontrada) {
            listaId = null;
            listaDisponible = false;

            console.error(
                "[ALFERZA] No se encontró la lista configurada.",
                {
                    listaBuscada: CONFIG.lista.nombre,
                    sitiosConsultados: diagnostico,
                    errores: erroresSitios
                }
            );

            // No bloqueamos la búsqueda de PDF por la lista.
            // El chatbot podrá responder, pero avisará si no puede guardar.
            return false;
        }

        listaId = encontrada.id;
        listaDisponible = true;

        console.info("[ALFERZA] Lista localizada:", {
            nombre: encontrada.displayName,
            id: encontrada.id,
            sitio: nombreSitioLista
        });

        const columnas = await obtenerTodasLasPaginas(
            "/sites/" + siteListaId +
            "/lists/" + listaId +
            "/columns?$select=id,name,displayName,hidden"
        );

        columnasLista = {};

        columnas.forEach(function (c) {
            if (c.hidden) return;

            // Guardamos tanto el nombre visible como el nombre interno.
            if (c.displayName) {
                columnasLista[nombreInterno(c.displayName)] = c.name;
            }

            if (c.name) {
                columnasLista[nombreInterno(c.name)] = c.name;
            }
        });

        console.info("[ALFERZA] Columnas de la lista:", columnasLista);

        return true;
    }

    function columna(nombre) {
        return columnasLista[nombreInterno(nombre)] || null;
    }

    /* =====================================================
       5. PDF.JS
    ===================================================== */

    async function cargarPdfJs() {
        if (pdfJsCargado && window.pdfjsLib) return;

        if (!window.pdfjsLib) {
            await new Promise(function (resolve, reject) {
                const script = document.createElement("script");

                script.src = CONFIG.pdfJsUrl;
                script.async = true;
                script.onload = resolve;
                script.onerror = function () {
                    reject(new Error(
                        "No se pudo cargar PDF.js desde el CDN."
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
       6. LOCALIZAR Y LEER LOS PDF
    ===================================================== */

    async function listarArchivosCarpeta() {
        const driveId = driveDocumentosId;
        const carpeta = CONFIG.documentos.carpeta;
        let carpetaItem = null;

        // Primero buscamos la carpeta por ruta.
        try {
            carpetaItem = await graphFetch(
                "/drives/" + driveId +
                "/root:/" + encodeURIComponent(carpeta) +
                "?$select=id,name,folder,webUrl"
            );
        } catch (error) {
            console.warn(
                "[ALFERZA] No se encontró la carpeta por ruta; " +
                "se intentará buscar en la raíz.",
                error.message
            );
        }

        // Si la ruta no funcionó, buscar por nombre en la raíz.
        if (!carpetaItem?.id) {
            const raiz = await obtenerTodasLasPaginas(
                "/drives/" + driveId +
                "/root/children?$select=id,name,folder,file,webUrl"
            );

            carpetaItem = raiz.find(function (item) {
                return item.folder &&
                    normalizar(item.name) === normalizar(carpeta);
            });
        }

        if (!carpetaItem?.id) {
            throw new Error(
                "No encontré la carpeta '" + carpeta +
                "' en la biblioteca '" + (driveDocumentosId || "") +
                "'. Comprueba el nombre y la ubicación."
            );
        }

        const archivos = [];
        const pendientes = [carpetaItem.id];
        const visitadas = new Set();

        while (pendientes.length) {
            const idCarpeta = pendientes.shift();

            if (visitadas.has(idCarpeta)) continue;
            visitadas.add(idCarpeta);

            const elementos = await obtenerTodasLasPaginas(
                "/drives/" + driveId +
                "/items/" + idCarpeta +
                "/children?$select=id,name,file,folder,webUrl,size"
            );

            for (const item of elementos) {
                if (item.folder) {
                    pendientes.push(item.id);
                } else if (item.file && /\.pdf$/i.test(item.name || "")) {
                    archivos.push(item);
                }
            }
        }

        console.info(
            "[ALFERZA] PDF encontrados:",
            archivos.map(a => a.name)
        );

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
                .map(item => item.str || "")
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

        if (actual.length) fragmentos.push(actual.join(" "));

        return fragmentos.filter(t => t.trim().length > 35);
    }

    async function indexarDocumentos(forzar) {
        if (
            !forzar &&
            documentosIndexados.length &&
            Date.now() - ultimaCargaDocumentos < CONFIG.cacheMs
        ) {
            return documentosIndexados;
        }

        if (indexacionEnCurso) return indexacionEnCurso;

        indexacionEnCurso = (async function () {
            await cargarPdfJs();

            const archivos = await listarArchivosCarpeta();

            if (!archivos.length) {
                throw new Error(
                    "No hay archivos PDF en la carpeta " +
                    CONFIG.documentos.carpeta + "."
                );
            }

            const indice = [];
            const errores = [];

            for (const archivo of archivos) {
                try {
                    const paginas = await extraerTextoPdf(archivo);

                    for (const pagina of paginas) {
                        dividirEnFragmentos(pagina.texto, 900)
                            .forEach(function (fragmento) {
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
       7. BUSCADOR DE CONTENIDO
    ===================================================== */

    const PALABRAS_IGNORADAS = new Set([
        "como", "que", "para", "por", "con", "sin", "una", "uno",
        "unos", "unas", "del", "las", "los", "este", "esta",
        "esto", "ese", "esa", "eso", "hay", "ser", "hacer",
        "tengo", "tiene", "puedo", "puede", "quiero", "necesito",
        "donde", "cuando", "cual", "cuales", "quien", "porque",
        "sobre", "entre", "desde", "hasta", "muy", "mas",
        "mi", "mis", "tu", "tus", "su", "sus", "el", "la",
        "un", "y", "o", "a", "de", "en", "es", "al", "me",
        "te", "se", "lo", "le", "ya", "si", "no", "hola",
        "buenos", "buenas"
    ]);

    function palabrasClave(texto) {
        return [...new Set(
            normalizar(texto)
                .split(/\s+/)
                .filter(p =>
                    p.length >= 2 && !PALABRAS_IGNORADAS.has(p)
                )
        )];
    }

    function buscarFragmentos(pregunta) {
        const terminos = palabrasClave(pregunta);
        if (!terminos.length) return [];

        const resultados = documentosIndexados
            .map(function (fragmento) {
                let puntuacion = 0;

                for (const termino of terminos) {
                    if (fragmento.normalizado.includes(termino)) {
                        puntuacion += termino.length >= 7 ? 3 : 1;
                    }
                }

                return {
                    ...fragmento,
                    puntuacion: puntuacion
                };
            })
            .filter(fragmento => fragmento.puntuacion > 0)
            .sort((a, b) => b.puntuacion - a.puntuacion);

        const seleccionados = [];
        const paginasUsadas = new Set();

        for (const resultado of resultados) {
            const clave = resultado.archivo + "|" + resultado.pagina;

            if (paginasUsadas.has(clave)) continue;

            paginasUsadas.add(clave);
            seleccionados.push(resultado);

            if (seleccionados.length >= CONFIG.maxResultados) break;
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
                    "procedimiento. También puedes indicar más detalles.",
                mode: "sin-resultados",
                fuentes: []
            };
        }

        const bloques = resultados.map(function (r, i) {
            return (
                (i + 1) + ". " + r.archivo +
                " (página " + r.pagina + ")\n" +
                r.texto
            );
        });

        return {
            text:
                "Encontré información relacionada en tus procedimientos " +
                "de TI:\n\n" + bloques.join("\n\n") +
                "\n\nEstos son fragmentos de los documentos originales. " +
                "Puedes pedirme que busque otro término o procedimiento.",
            mode: "sharepoint",
            fuentes: resultados.map(r => ({
                titulo: r.archivo,
                pagina: r.pagina,
                url: r.webUrl
            }))
        };
    }

    /* =====================================================
       8. HISTORIAL Y CONTEXTO
    ===================================================== */

    function agregarHistorial(role, content) {
        historial.push({
            role: role,
            content: String(content || ""),
            timestamp: Date.now()
        });

        if (historial.length > 100) historial = historial.slice(-100);
        guardarHistorialLocal();
    }

    function historialContexto() {
        return historial.slice(-CONFIG.maxMensajesContexto).map(m => ({
            role: m.role,
            content: m.content
        }));
    }

    function enriquecerPreguntaConContexto(pregunta) {
        const actual = normalizar(pregunta);

        const seguimiento =
            /^(y |y si |y como |eso |ese |esa |lo mismo|tambien|por que |donde |como hago |explica|amplia|detalla|continua)/i
                .test(actual) ||
            actual.split(/\s+/).length <= 3;

        if (!seguimiento) return pregunta;

        const ultimaPregunta = [...historial]
            .reverse()
            .find(m => m.role === "user");

        return ultimaPregunta
            ? ultimaPregunta.content + " " + pregunta
            : pregunta;
    }

    /* =====================================================
       9. HISTORIAL REMOTO EN MICROSOFT LISTS
    ===================================================== */

    async function cargarHistorialRemoto() {
        if (!listaDisponible || !listaId || !siteListaId) {
            console.warn(
                "[ALFERZA] Historial remoto desactivado: la lista no está localizada."
            );
            return;
        }

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
                "[ALFERZA] La lista no tiene todas las columnas necesarias " +
                "para recuperar el historial. Se usará el historial local."
            );
            return;
        }

        const filtro =
            "fields/" + colConversation + " eq '" +
            conversationId.replace(/'/g, "''") +
            "' and fields/" + colCorreo + " eq '" +
            usuario.correo.replace(/'/g, "''") + "'";

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
            encodeURIComponent(seleccion) +
            ")&$filter=" + encodeURIComponent(filtro) +
            "&$top=100";

        try {
            const registros = await obtenerTodasLasPaginas(url);
            const remoto = [];

            registros.forEach(function (registro) {
                const campos = registro.fields || {};
                const pregunta = campos[colPregunta];
                const respuesta = campos[colRespuesta];
                const timestamp = Date.parse(campos[colFecha] || "") || 0;

                if (pregunta) {
                    remoto.push({
                        role: "user",
                        content: String(pregunta),
                        timestamp: timestamp
                    });
                }

                if (respuesta) {
                    remoto.push({
                        role: "assistant",
                        content: String(respuesta),
                        timestamp: timestamp
                    });
                }
            });

            remoto.sort((a, b) => a.timestamp - b.timestamp);

            if (remoto.length) {
                historial = remoto.slice(-100);
                guardarHistorialLocal();
            }

            console.info(
                "[ALFERZA] Mensajes remotos recuperados:",
                remoto.length
            );
        } catch (error) {
            // No descargar todos los registros: podrían ser de otros usuarios.
            console.warn(
                "[ALFERZA] No se pudo recuperar el historial remoto:",
                error
            );
        }
    }

    async function guardarIntercambio(pregunta, respuesta, fuentes, tema) {
        if (!listaDisponible || !listaId || !siteListaId) {
            throw new Error(
                "El chatbot respondió, pero no se encontró la lista " +
                CONFIG.lista.nombre +
                ". Revisa en la consola los sitios y nombres de listas encontrados."
            );
        }

        const campos = {};

        const colTitulo = columna("Title") || columna("Título");
        const colConversation = columna("ConversationId");
        const colCorreo = columna("UsuarioCorreo");
        const colPregunta = columna("MensajeUsuario");
        const colRespuesta = columna("RespuestaBot");
        const colFecha = columna("Fecha");
        const colDocumento = columna("DocumentoOrigen");
        const colTema = columna("Tema");

        if (colTitulo) campos[colTitulo] = pregunta.slice(0, 255);
        if (colConversation) campos[colConversation] = conversationId;
        if (colCorreo) campos[colCorreo] = usuario.correo;
        if (colPregunta) campos[colPregunta] = pregunta;
        if (colRespuesta) campos[colRespuesta] = respuesta;

        if (colFecha) {
            campos[colFecha] = new Date().toISOString();
        }

        if (colDocumento) {
            campos[colDocumento] = (fuentes || [])
                .map(f =>
                    f.titulo + (f.pagina ? " (p. " + f.pagina + ")" : "")
                )
                .join("; ")
                .slice(0, 255);
        }

        if (colTema) {
            campos[colTema] = String(tema || "Soporte TI").slice(0, 255);
        }

        const faltantes = [];

        if (!colConversation) faltantes.push("ConversationId");
        if (!colCorreo) faltantes.push("UsuarioCorreo");
        if (!colPregunta) faltantes.push("MensajeUsuario");
        if (!colRespuesta) faltantes.push("RespuestaBot");

        if (faltantes.length) {
            throw new Error(
                "Faltan columnas en la lista: " + faltantes.join(", ") +
                ". Revisa los nombres de las columnas en Microsoft Lists."
            );
        }

        await graphFetch(
            "/sites/" + siteListaId +
            "/lists/" + listaId + "/items",
            {
                method: "POST",
                body: JSON.stringify({ fields: campos })
            }
        );

        console.info("[ALFERZA] Intercambio guardado en Microsoft Lists.");
    }

    /* =====================================================
       10. CACHÉ DE RESPUESTAS
    ===================================================== */

    function obtenerRespuestaCache(pregunta) {
        const clave = normalizar(pregunta);
        const entrada = cache[clave];

        if (!entrada || typeof entrada.text !== "string") return null;

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
        if (!respuesta?.text) return;

        cache[normalizar(pregunta)] = {
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
       11. ESTILOS, MENSAJES Y DESPLAZAMIENTO
    ===================================================== */

    function instalarEstilos() {
        if (document.getElementById("alferza-chat-dynamic-styles")) return;

        const estilos = document.createElement("style");
        estilos.id = "alferza-chat-dynamic-styles";

        estilos.textContent = `
            #typingIndicator {
                transition: opacity .2s ease;
            }

            #typingIndicator:not([hidden]) {
                animation: alferzaTypingPulse 1.2s ease-in-out infinite;
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

            @media (prefers-reduced-motion: reduce) {
                #typingIndicator {
                    animation: none !important;
                    transition: none !important;
                }
            }
        `;

        document.head.appendChild(estilos);
    }

    function desplazarAlFinal(inmediato) {
        const desplazar = function () {
            elements.messages.scrollTop = elements.messages.scrollHeight;
        };

        if (inmediato) desplazar();

        requestAnimationFrame(function () {
            desplazar();
            requestAnimationFrame(desplazar);
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
        fila.className = "message-row " + tipo;

        const contenido = document.createElement("div");
        contenido.className = "message-content";

        const nombre = document.createElement("span");
        nombre.className = "message-name";
        nombre.textContent = tipo === "user" ? "Tú" : "Asistente ALFERZA";

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

        elements.suggestions.forEach(boton => {
            boton.disabled = visible;
        });

        if (visible) desplazarAlFinal();
    }

    function actualizarContador() {
        const cantidad = elements.input.value.length;

        elements.characterCount.textContent =
            cantidad + "/" + CONFIG.maxCaracteres;

        elements.input.style.height = "auto";
        elements.input.style.height =
            Math.min(elements.input.scrollHeight, 120) + "px";
    }

    function mostrarFuentesComoTexto(fuentes) {
        if (!fuentes?.length) return "";

        return "\n\nDocumentos consultados:\n" +
            fuentes.map(function (fuente) {
                return "• " + fuente.titulo +
                    (fuente.pagina ? " (página " + fuente.pagina + ")" : "");
            }).join("\n");
    }

    /* =====================================================
       12. CONSULTA Y ENVÍO
    ===================================================== */

    async function consultarDocumentos(mensaje) {
        const pregunta = enriquecerPreguntaConContexto(mensaje);

        await indexarDocumentos(false);

        const resultado = respuestaDesdeDocumentos(pregunta);

        return {
            text: resultado.text + mostrarFuentesComoTexto(resultado.fuentes),
            mode: resultado.mode,
            fuentes: resultado.fuentes
        };
    }

    async function enviarMensaje(mensaje) {
        const texto = String(mensaje || "").trim();

        if (!texto || procesandoMensaje) return;

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
                "El asistente todavía está conectándose con Microsoft 365. " +
                "Espera unos segundos e inténtalo nuevamente."
            );
            return;
        }

        procesandoMensaje = true;
        agregarMensaje("user", texto);
        agregarHistorial("user", texto);

        elements.input.value = "";
        actualizarContador();
        mostrarEscritura(true);

        try {
            let respuesta = obtenerRespuestaCache(texto);

            if (!respuesta) {
                respuesta = await consultarDocumentos(texto);
                guardarRespuestaCache(texto, respuesta);
            }

            agregarMensaje("bot", respuesta.text);
            agregarHistorial("assistant", respuesta.text);

            try {
                await guardarIntercambio(
                    texto,
                    respuesta.text,
                    respuesta.fuentes,
                    respuesta.mode
                );
            } catch (errorGuardar) {
                console.error(
                    "[ALFERZA] Error guardando en Microsoft Lists:",
                    errorGuardar
                );

                agregarMensaje(
                    "bot",
                    "Aviso: la respuesta se generó, pero no se pudo guardar " +
                    "en Microsoft Lists.\n\n" +
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
                    CONFIG.documentos.carpeta +
                    " y que los PDF permitan seleccionar texto.";
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
       13. EVENTOS
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
            if (boton.dataset.question) {
                enviarMensaje(boton.dataset.question);
            }
        });
    });

    elements.clearButton.addEventListener("click", function () {
        if (procesandoMensaje) return;

        conversationId = generarId();
        historial = [];
        cache = {};

        try {
            sessionStorage.setItem(STORAGE_KEY, conversationId);
        } catch (_) {}

        guardarHistorialLocal();
        guardarCache();
        elements.messages.replaceChildren();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "He iniciado una conversación nueva. Puedo buscar información " +
            "en los procedimientos internos. ¿Qué necesitas resolver hoy?"
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

            // Si no encuentra la lista, todavía permite buscar en los PDF.
            await prepararSharePoint();

            if (listaDisponible) {
                await cargarHistorialRemoto();
            } else {
                agregarMensaje(
                    "bot",
                    "Aviso: no pude localizar la lista " +
                    CONFIG.lista.nombre +
                    ". Podrás consultar los procedimientos, pero el guardado " +
                    "remoto estará desactivado hasta corregir la ubicación de la lista."
                );
            }

            inicializado = true;

            agregarMensaje(
                "bot",
                "Conexión con Microsoft 365 establecida. " +
                "Ya puedes hacer consultas sobre los procedimientos " +
                "disponibles en SharePoint."
            );

            console.info("[ALFERZA CHATBOT] Inicialización completada.", {
                conversationId: conversationId,
                usuario: usuario.correo,
                listaDisponible: listaDisponible,
                sitioLista: nombreSitioLista || "No localizado",
                mensajesLocales: historial.length,
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
                "\n\nAbre F12 > Console para ver el detalle técnico."
            );
        } finally {
            mostrarEscritura(false);
            actualizarContador();
            desplazarAlFinal();
        }
    }

    iniciarChatbot();
})();