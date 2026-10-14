"use strict";

/* =========================================================
   ALFERZA LIVE OFFICE
   ASISTENTE DE SOPORTE TI
   Conexión con backend Node.js + Ollama + SharePoint
========================================================= */

(function () {
    const API_URL = "http://localhost:3000/api/chat";
    const STORAGE_KEY = "alferza_chatbot_conversation_id";

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
            "[ALFERZA CHATBOT] No se encontraron todos los elementos del chat."
        );
        return;
    }

    let procesandoMensaje = false;
    let conversationId = obtenerConversationId();

    /* =====================================================
       1. IDENTIFICADOR DE CONVERSACIÓN
    ===================================================== */

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
            console.warn(
                "[ALFERZA CHATBOT] No se pudo acceder a sessionStorage."
            );

            return generarId();
        }
    }

    function crearNuevaConversacion() {
        conversationId = generarId();

        try {
            sessionStorage.setItem(STORAGE_KEY, conversationId);
        } catch (error) {
            console.warn(
                "[ALFERZA CHATBOT] No se pudo guardar el identificador."
            );
        }
    }

    /* =====================================================
       2. MENSAJES DE LA INTERFAZ
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

        // Nunca interpretar la respuesta de la IA como HTML.
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

    function desplazarAlFinal() {
        elements.messages.scrollTop =
            elements.messages.scrollHeight;
    }

    /* =====================================================
       3. ESTADO DE LA INTERFAZ
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
            `${cantidad}/1000`;

        elements.input.style.height = "auto";
        elements.input.style.height =
            `${Math.min(elements.input.scrollHeight, 120)}px`;
    }

    /* =====================================================
       4. INTERPRETAR LA RESPUESTA DEL BACKEND
    ===================================================== */

    function obtenerTextoRespuesta(datos) {
        return (
            datos.respuesta ||
            datos.answer ||
            datos.response ||
            datos.mensaje ||
            datos.message ||
            ""
        );
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
                    ""
                );
            }

            return "";
        }).filter(Boolean);

        if (nombres.length === 0) {
            return "";
        }

        return "\n\nFuentes consultadas:\n" +
            [...new Set(nombres)].join("\n");
    }

    /* =====================================================
       5. CONSULTAR LA IA DEL BACKEND
    ===================================================== */

    async function consultarBackend(mensaje) {
        const respuestaHTTP = await fetch(API_URL, {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                conversationId: conversationId,
                message: mensaje
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
                "El servidor respondió, pero no devolvió el texto de la respuesta."
            );
        }

        // Permite conservar el ID que devuelve el servidor.
        if (datos.conversationId) {
            conversationId = datos.conversationId;

            try {
                sessionStorage.setItem(
                    STORAGE_KEY,
                    conversationId
                );
            } catch (error) {
                console.warn(
                    "[ALFERZA CHATBOT] No se pudo guardar la sesión."
                );
            }
        }

        return texto + obtenerFuentes(datos);
    }

    /* =====================================================
       6. PROCESAR Y ENVIAR MENSAJES
    ===================================================== */

    async function enviarMensaje(mensaje) {
        const texto = String(mensaje || "").trim();

        if (!texto || procesandoMensaje) {
            return;
        }

        if (texto.length > 1000) {
            agregarMensaje(
                "bot",
                "La consulta supera el límite de 1000 caracteres."
            );
            return;
        }

        procesandoMensaje = true;

        agregarMensaje("user", texto);

        elements.input.value = "";
        actualizarContador();
        mostrarEscritura(true);

        try {
            const respuesta = await consultarBackend(texto);

            agregarMensaje("bot", respuesta);

            console.info(
                "[ALFERZA CHATBOT] Respuesta recibida correctamente."
            );
        } catch (error) {
            console.error(
                "[ALFERZA CHATBOT] Error de conexión:",
                error
            );

            agregarMensaje(
                "bot",
                "No pude comunicarme con el servicio de inteligencia artificial.\n\n" +
                "Verifica que:\n" +
                "1. El backend Node.js esté ejecutándose.\n" +
                "2. El servicio de Ollama esté disponible.\n" +
                "3. La dirección del backend sea accesible desde este navegador.\n\n" +
                "Tu consulta no se ha confirmado como procesada. " +
                "Corrige la conexión e inténtalo nuevamente."
            );
        } finally {
            mostrarEscritura(false);

            procesandoMensaje = false;
            elements.input.disabled = false;
            elements.sendButton.disabled = false;

            actualizarContador();
            elements.input.focus();
        }
    }

    /* =====================================================
       7. EVENTOS DEL CHAT
    ===================================================== */

    elements.form.addEventListener("submit", function (evento) {
        evento.preventDefault();
        enviarMensaje(elements.input.value);
    });

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
       8. NUEVA CONVERSACIÓN
       No elimina el historial de SQLite.
    ===================================================== */

    elements.clearButton.addEventListener("click", function () {
        if (procesandoMensaje) {
            return;
        }

        crearNuevaConversacion();

        elements.messages.replaceChildren();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "He iniciado una conversación nueva. " +
            "Puedo ayudarte con los procedimientos de TI disponibles.\n\n" +
            "¿Qué necesitas resolver hoy?"
        );

        elements.input.value = "";
        actualizarContador();
        elements.input.focus();
    });

    /* =====================================================
       9. INICIALIZACIÓN
    ===================================================== */

    function iniciarChatbot() {
        actualizarContador();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "Estoy conectado al servicio de soporte cuando el backend está disponible. " +
            "Puedo consultar los procedimientos de TI indexados desde SharePoint.\n\n" +
            "Cuéntame qué necesitas resolver."
        );

        console.info(
            "[ALFERZA CHATBOT] Interfaz inicializada.",
            { conversationId: conversationId }
        );
    }

    iniciarChatbot();
})();