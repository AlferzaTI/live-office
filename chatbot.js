
"use strict";

/* =========================================================
   ALFERZA LIVE OFFICE
   ASISTENTE DE SOPORTE TI
   Respuestas completas + caché + historial
   Desplazamiento automático mejorado
========================================================= */

(function () {
    const API_URL = "http://localhost:3000/api/chat";

    const STORAGE_KEY = "alferza_chatbot_conversation_id";
    const HISTORY_PREFIX = "alferza_chat_history_";
    const CACHE_PREFIX = "alferza_chat_cache_";

    const MAX_CARACTERES = 1000;
    const MAX_HISTORIAL_ENVIADO = 6;
    const MAX_ENTRADAS_CACHE = 100;

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

    let procesandoMensaje = false;
    let conversationId = obtenerConversationId();
    let historial = cargarHistorial();
    let cache = cargarCache();

    /* =====================================================
       1. ALMACENAMIENTO DE SESIÓN
    ===================================================== */

    function leerStorage(clave, valorPredeterminado) {
        try {
            const valor = sessionStorage.getItem(clave);

            return valor
                ? JSON.parse(valor)
                : valorPredeterminado;
        } catch (error) {
            console.warn(
                "[ALFERZA] No se pudo leer el almacenamiento."
            );

            return valorPredeterminado;
        }
    }

    function guardarStorage(clave, valor) {
        try {
            sessionStorage.setItem(
                clave,
                JSON.stringify(valor)
            );

            return true;
        } catch (error) {
            console.warn(
                "[ALFERZA] No se pudo guardar la sesión."
            );

            return false;
        }
    }

    function generarId() {
        if (
            window.crypto &&
            typeof window.crypto.randomUUID === "function"
        ) {
            return window.crypto.randomUUID();
        }

        return "chat-" +
            Date.now() +
            "-" +
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
        } catch (error) {
            return generarId();
        }
    }

    function claveHistorial() {
        return HISTORY_PREFIX + conversationId;
    }

    function claveCache() {
        return CACHE_PREFIX + conversationId;
    }

    function cargarHistorial() {
        const datos = leerStorage(
            HISTORY_PREFIX + conversationId,
            []
        );

        return Array.isArray(datos) ? datos : [];
    }

    function guardarHistorial() {
        guardarStorage(claveHistorial(), historial);
    }

    function cargarCache() {
        const datos = leerStorage(
            CACHE_PREFIX + conversationId,
            {}
        );

        return datos &&
            typeof datos === "object" &&
            !Array.isArray(datos)
                ? datos
                : {};
    }

    function guardarCache() {
        guardarStorage(claveCache(), cache);
    }

    /* =====================================================
       2. CLASIFICACIÓN DE CONSULTAS ACTUALES
    ===================================================== */

    function necesitaInternet(pregunta) {
        return /\b(clima|tiempo|temperatura|pronóstico|pronostico|lluvia|llover|llueve|humedad|viento|meteorológico|meteorologica|meteorología|noticias|actualmente|hoy|ahora|último|última|ultimos|ultimas|cotización|cotizacion|precio actual|tipo de cambio|resultados en vivo|versión actual|version actual|este año|este mes|202[6-9])\b/i.test(
            pregunta
        );
    }

    function normalizarPregunta(pregunta) {
        return String(pregunta || "")
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[¿?¡!.,;:]/g, "")
            .replace(/\s+/g, " ")
            .trim();
    }

    function esCacheable(pregunta) {
        return !necesitaInternet(pregunta);
    }

    /* =====================================================
       3. CACHÉ DE RESPUESTAS
    ===================================================== */

    function obtenerRespuestaCache(pregunta) {
        if (!esCacheable(pregunta)) {
            return null;
        }

        const clave = normalizarPregunta(pregunta);
        const entrada = cache[clave];

        if (!entrada || typeof entrada.text !== "string") {
            return null;
        }

        return entrada;
    }

    function guardarRespuestaCache(pregunta, respuesta) {
        if (
            !esCacheable(pregunta) ||
            !respuesta ||
            !respuesta.text ||
            respuesta.mode === "error"
        ) {
            return;
        }

        const clave = normalizarPregunta(pregunta);

        cache[clave] = {
            text: respuesta.text,
            mode: respuesta.mode || "general",
            savedAt: Date.now()
        };

        const claves = Object.keys(cache);

        while (claves.length > MAX_ENTRADAS_CACHE) {
            const masAntigua = claves.shift();
            delete cache[masAntigua];
        }

        guardarCache();
    }

    /* =====================================================
       4. HISTORIAL DE CONVERSACIÓN
    ===================================================== */

    function registrarHistorial(role, content) {
        historial.push({
            role,
            content: String(content || ""),
            timestamp: Date.now()
        });

        guardarHistorial();
    }

    function obtenerHistorialParaBackend() {
        return historial
            .slice(-MAX_HISTORIAL_ENVIADO)
            .map(function (mensaje) {
                return {
                    role: mensaje.role,
                    content: mensaje.content
                };
            });
    }

    /* =====================================================
       5. ESTILOS DINÁMICOS
       Sin límite de líneas.
    ===================================================== */

    function instalarEstilos() {
        if (
            document.getElementById(
                "alferza-chat-dynamic-styles"
            )
        ) {
            return;
        }

        const estilos = document.createElement("style");
        estilos.id = "alferza-chat-dynamic-styles";

        estilos.textContent = `
            #typingIndicator {
                transition: opacity 0.2s ease;
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
                animation-delay: 0.15s;
            }

            .alferza-typing-dots span:nth-child(3) {
                animation-delay: 0.3s;
            }

            /* Mostrar siempre la respuesta completa. */
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

            .message-row.bot .message-avatar {
                transition: transform 0.2s ease;
            }

            .message-row.bot:hover .message-avatar {
                transform: scale(1.05);
            }

            @keyframes alferzaTypingPulse {
                0%, 100% { opacity: 0.65; }
                50% { opacity: 1; }
            }

            @keyframes alferzaDotBounce {
                0%, 60%, 100% {
                    transform: translateY(0);
                    opacity: 0.45;
                }
                30% {
                    transform: translateY(-4px);
                    opacity: 1;
                }
            }

            @media (prefers-reduced-motion: reduce) {
                #typingIndicator,
                .alferza-typing-dots span,
                .message-row.bot .message-avatar {
                    animation: none !important;
                    transition: none !important;
                }
            }
        `;

        document.head.appendChild(estilos);

        if (
            !elements.typing.querySelector(
                ".alferza-typing-dots"
            )
        ) {
            const puntos = document.createElement("span");
            puntos.className = "alferza-typing-dots";
            puntos.setAttribute("aria-hidden", "true");

            for (let i = 0; i < 3; i++) {
                puntos.appendChild(
                    document.createElement("span")
                );
            }

            elements.typing.appendChild(puntos);
        }

        elements.typing.setAttribute("role", "status");
        elements.typing.setAttribute(
            "aria-label",
            "ALFERZA está generando una respuesta"
        );
    }

    /* =====================================================
       6. DESPLAZAMIENTO AUTOMÁTICO
       Mantiene visible el último mensaje.
    ===================================================== */

    function desplazarAlFinal(inmediato = false) {
        const desplazar = function () {
            elements.messages.scrollTop =
                elements.messages.scrollHeight;

            // Mantener visible el indicador de escritura
            // si está dentro del área de mensajes.
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
       7. MENSAJES DE LA INTERFAZ
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

        // Mostrar el contenido como texto seguro.
        // No interpretar la respuesta como HTML.
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

    /* =====================================================
       8. ESTADO DE LA INTERFAZ
    ===================================================== */

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
            `${cantidad}/${MAX_CARACTERES}`;

        elements.input.style.height = "auto";

        elements.input.style.height =
            `${Math.min(elements.input.scrollHeight, 120)}px`;
    }

    /* =====================================================
       9. RESPUESTA DEL BACKEND
    ===================================================== */

    function obtenerTextoRespuesta(datos) {
        const posiblesRespuestas = [
            datos.respuesta,
            datos.answer,
            datos.response,
            datos.mensaje,
            datos.message
        ];

        for (const respuesta of posiblesRespuestas) {
            if (
                typeof respuesta === "string" &&
                respuesta.trim()
            ) {
                return respuesta.trim();
            }
        }

        return "";
    }

    function obtenerFuentes(datos) {
        const fuentes =
            datos.fuentes ||
            datos.sources ||
            datos.fuentesConsultadas ||
            [];

        if (!Array.isArray(fuentes) || fuentes.length === 0) {
            return "";
        }

        const nombres = fuentes.map(function (fuente) {
            if (typeof fuente === "string") {
                return fuente;
            }

            if (fuente && typeof fuente === "object") {
                return (
                    fuente.titulo ||
                    fuente.nombre ||
                    fuente.archivo ||
                    fuente.fuente ||
                    fuente.source ||
                    fuente.title ||
                    ""
                );
            }

            return "";
        }).filter(Boolean);

        if (nombres.length === 0) {
            return "";
        }

        return "\n\nFuentes: " +
            [...new Set(nombres)].join(", ");
    }

    /* =====================================================
       10. CONSULTAR EL BACKEND
    ===================================================== */

    async function consultarBackend(mensaje) {
        const respuestaHTTP = await fetch(API_URL, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                conversationId: conversationId,
                message: mensaje,
                history: obtenerHistorialParaBackend()
            })
        });

        let datos;

        try {
            datos = await respuestaHTTP.json();
        } catch (error) {
            throw new Error(
                "El servidor devolvió una respuesta que no es JSON válido."
            );
        }

        if (!respuestaHTTP.ok || datos.success === false) {
            throw new Error(
                datos.error ||
                datos.message ||
                `Error del servidor: HTTP ${respuestaHTTP.status}`
            );
        }

        const texto = obtenerTextoRespuesta(datos);

        if (!texto) {
            throw new Error(
                "El servidor no devolvió el texto de la respuesta."
            );
        }

        if (datos.conversationId) {
            conversationId = datos.conversationId;

            try {
                sessionStorage.setItem(
                    STORAGE_KEY,
                    conversationId
                );
            } catch (error) {
                console.warn(
                    "[ALFERZA] No se pudo guardar el identificador."
                );
            }
        }

        return {
            text: texto + obtenerFuentes(datos),
            mode: datos.mode || datos.modo || "general"
        };
    }

    /* =====================================================
       11. ENVIAR MENSAJES
       Sin recortar respuestas de caché ni del backend.
    ===================================================== */

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
        desplazarAlFinal();

        try {
            const respuestaGuardada =
                obtenerRespuestaCache(texto);

            if (respuestaGuardada) {
                console.info(
                    "[ALFERZA] Caché utilizada."
                );

                agregarMensaje(
                    "bot",
                    respuestaGuardada.text
                );

                registrarHistorial(
                    "assistant",
                    respuestaGuardada.text
                );

                return;
            }

            console.info(
                "[ALFERZA] Enviando consulta al backend."
            );

            const respuesta = await consultarBackend(texto);

            agregarMensaje(
                "bot",
                respuesta.text
            );

            registrarHistorial(
                "assistant",
                respuesta.text
            );

            guardarRespuestaCache(texto, respuesta);

            console.info(
                "[ALFERZA] Respuesta recibida correctamente."
            );
        } catch (error) {
            console.error(
                "[ALFERZA CHATBOT] Error:",
                error
            );

            agregarMensaje(
                "bot",
                "No pude obtener una respuesta.\n\n" +
                "Verifica que el backend Node.js esté funcionando. " +
                "Si el problema continúa, revisa la terminal del backend."
            );
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

    elements.form.addEventListener(
        "submit",
        function (evento) {
            evento.preventDefault();
            enviarMensaje(elements.input.value);
        }
    );

    elements.input.addEventListener(
        "input",
        actualizarContador
    );

    elements.input.addEventListener(
        "keydown",
        function (evento) {
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
        }
    );

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

    elements.clearButton.addEventListener(
        "click",
        function () {
            if (procesandoMensaje) {
                return;
            }

            conversationId = generarId();
            historial = [];
            cache = {};

            try {
                sessionStorage.setItem(
                    STORAGE_KEY,
                    conversationId
                );
            } catch (error) {
                console.warn(
                    "[ALFERZA] No se pudo guardar la nueva conversación."
                );
            }

            guardarHistorial();
            guardarCache();

            elements.messages.replaceChildren();

            agregarMensaje(
                "bot",
                "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
                "He iniciado una conversación nueva. " +
                "¿Qué necesitas resolver hoy?"
            );

            elements.input.value = "";
            actualizarContador();
            desplazarAlFinal();

            elements.input.focus();

            console.info(
                "[ALFERZA] Nueva conversación iniciada."
            );
        }
    );

    /* =====================================================
       14. INICIALIZACIÓN
    ===================================================== */

    function iniciarChatbot() {
        instalarEstilos();
        actualizarContador();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "Puedo ayudarte con tus consultas y los procedimientos " +
            "internos disponibles."
        );

        desplazarAlFinal();

        console.info(
            "[ALFERZA CHATBOT] Inicializado.",
            {
                conversationId: conversationId,
                mensajesGuardados: historial.length,
                respuestasEnCache: Object.keys(cache).length
            }
        );
    }

    iniciarChatbot();
})();
