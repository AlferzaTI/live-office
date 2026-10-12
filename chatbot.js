
"use strict";

/* =========================================================
   ALFERZA LIVE OFFICE
   ASISTENTE DE SOPORTE TI
   Versión inicial: reglas y base de conocimiento local
========================================================= */

(function () {
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
        !elements.sendButton
    ) {
        console.error(
            "[ALFERZA CHATBOT] No se encontraron los elementos del chat."
        );
        return;
    }

    let procesandoMensaje = false;

    /* =====================================================
       1. BASE DE CONOCIMIENTO
       Agrega aquí procedimientos aprobados por TI.
    ===================================================== */

    const BASE_CONOCIMIENTO = [
        {
            id: "saludo",
            categoria: "General",
            palabras: [
                "hola",
                "buenos dias",
                "buenas tardes",
                "buenas noches",
                "que tal",
                "buenas",
                "hey"
            ],
            respuesta:
                "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
                "Puedo orientarte con problemas informáticos, " +
                "instalaciones, conexión a internet, Outlook y solicitudes de soporte.\n\n" +
                "Cuéntame qué necesitas resolver."
        },
        {
            id: "instalar_software",
            categoria: "Instalaciones",
            palabras: [
                "instalar",
                "instalacion",
                "instalame",
                "instalarme",
                "software",
                "programa",
                "aplicacion",
                "aplicativo",
                "descargar programa",
                "visual studio code",
                "vscode",
                "office",
                "anydesk",
                "teamviewer"
            ],
            respuesta:
                "SOLICITUD DE INSTALACIÓN\n\n" +
                "Para orientarte necesito conocer el programa que necesitas.\n\n" +
                "1. Indícame el nombre del software.\n" +
                "2. Explícame para qué lo utilizarás.\n" +
                "3. Señala si es una computadora de la empresa.\n\n" +
                "Si el programa requiere permisos de administrador, " +
                "debes solicitar la autorización correspondiente al área de TI.\n\n" +
                "Importante: este chatbot todavía no instala programas ni " +
                "comprueba las políticas de software de la empresa."
        },
        {
            id: "problema_red",
            categoria: "Redes",
            palabras: [
                "internet",
                "sin internet",
                "no tengo internet",
                "no hay internet",
                "wifi",
                "wi fi",
                "conexion",
                "conectividad",
                "red",
                "sin red",
                "no conecta",
                "no tengo conexion",
                "ethernet",
                "cable de red",
                "ip",
                "dns"
            ],
            respuesta:
                "DIAGNÓSTICO DE CONEXIÓN\n\n" +
                "Prueba estas verificaciones:\n\n" +
                "1. Comprueba si el Wi-Fi está conectado o si el cable de red está bien colocado.\n" +
                "2. Verifica si otras páginas o aplicaciones también presentan el problema.\n" +
                "3. Comprueba si otros compañeros tienen la misma incidencia.\n" +
                "4. Si utilizas una VPN corporativa, verifica si está conectada correctamente.\n\n" +
                "Si el problema continúa, solicita soporte de TI e indica tu área, " +
                "el nombre de tu equipo y desde cuándo ocurre.\n\n" +
                "No cambies la configuración de IP, DNS o proxy corporativo " +
                "sin autorización."
        },
        {
            id: "outlook",
            categoria: "Correo",
            palabras: [
                "outlook",
                "correo",
                "email",
                "e mail",
                "no puedo enviar",
                "no puedo recibir",
                "no llegan correos",
                "no envia correos",
                "buzon",
                "bandeja de entrada",
                "correo bloqueado",
                "correo lleno"
            ],
            respuesta:
                "SOPORTE DE OUTLOOK\n\n" +
                "Primero identifica qué sucede:\n\n" +
                "1. No puedes iniciar sesión.\n" +
                "2. No puedes enviar o recibir correos.\n" +
                "3. Outlook está lento o no abre.\n" +
                "4. Tu buzón parece estar lleno.\n\n" +
                "Prueba acceder a Outlook desde el navegador para comprobar " +
                "si el problema también ocurre allí.\n\n" +
                "Si sospechas que alguien accedió a tu cuenta, no compartas " +
                "contraseñas ni códigos de verificación. Comunícalo de inmediato " +
                "al área responsable de seguridad o TI."
        },
        {
            id: "impresora",
            categoria: "Hardware",
            palabras: [
                "impresora",
                "imprimir",
                "no imprime",
                "no puedo imprimir",
                "atasco de papel",
                "toner",
                "tinta",
                "impresion"
            ],
            respuesta:
                "DIAGNÓSTICO DE IMPRESORA\n\n" +
                "1. Comprueba que la impresora esté encendida.\n" +
                "2. Verifica si muestra algún error en su pantalla.\n" +
                "3. Revisa si tiene papel y consumibles disponibles.\n" +
                "4. Confirma que tu computadora tenga seleccionada la impresora correcta.\n" +
                "5. Comprueba si otros usuarios pueden imprimir.\n\n" +
                "Si continúa fallando, registra una solicitud de soporte e indica " +
                "el modelo o nombre de la impresora y el mensaje de error."
        },
        {
            id: "equipo_lento",
            categoria: "Equipos",
            palabras: [
                "computadora lenta",
                "pc lenta",
                "laptop lenta",
                "equipo lento",
                "muy lenta",
                "se congela",
                "congelada",
                "se cuelga",
                "no responde",
                "memoria ram",
                "disco lleno"
            ],
            respuesta:
                "DIAGNÓSTICO DE EQUIPO\n\n" +
                "1. Comprueba si hay aplicaciones que no estés utilizando.\n" +
                "2. Guarda tu trabajo antes de reiniciar el equipo.\n" +
                "3. Verifica si aparece algún mensaje de error o advertencia de almacenamiento.\n" +
                "4. Anota cuándo comenzó la lentitud y qué aplicaciones se ven afectadas.\n\n" +
                "Si el problema persiste, solicita una revisión de TI.\n\n" +
                "No elimines archivos corporativos ni instales programas de limpieza " +
                "sin autorización."
        },
        {
            id: "contrasena",
            categoria: "Accesos",
            palabras: [
                "contraseña",
                "contrasena",
                "password",
                "olvide mi clave",
                "olvide mi contraseña",
                "cambiar clave",
                "restablecer contraseña",
                "no puedo iniciar sesion",
                "bloqueo de cuenta",
                "cuenta bloqueada",
                "mfa",
                "autenticador",
                "authenticator"
            ],
            respuesta:
                "AYUDA CON TU CUENTA\n\n" +
                "Si no puedes acceder a tu cuenta corporativa:\n\n" +
                "1. Verifica que estés usando tu cuenta de trabajo correcta.\n" +
                "2. Comprueba si el mensaje de error indica una contraseña incorrecta, " +
                "una cuenta bloqueada o un problema de autenticación.\n" +
                "3. Utiliza el procedimiento oficial de recuperación de Microsoft " +
                "o contacta al área de TI.\n\n" +
                "Nunca compartas tu contraseña, PIN, códigos MFA ni códigos de recuperación " +
                "en este chat."
        },
        {
            id: "ticket",
            categoria: "Soporte",
            palabras: [
                "ticket",
                "tickets",
                "incidencia",
                "incidencias",
                "soporte",
                "solicitud de soporte",
                "reportar problema",
                "registrar problema",
                "crear ticket",
                "abrir ticket",
                "ayuda tecnica",
                "tecnico"
            ],
            respuesta:
                "SOLICITUD DE SOPORTE TI\n\n" +
                "Para preparar tu solicitud, proporciona estos datos:\n\n" +
                "• Nombre del problema.\n" +
                "• Equipo o aplicación afectada.\n" +
                "• Área donde trabajas.\n" +
                "• Descripción del error y cuándo comenzó.\n" +
                "• Impacto en tus actividades.\n\n" +
                "Esta versión todavía no crea registros en Microsoft Lists. " +
                "Por ahora, utiliza el canal de soporte establecido por tu empresa."
        },
        {
            id: "teams",
            categoria: "Aplicaciones",
            palabras: [
                "teams",
                "microsoft teams",
                "videollamada",
                "no escucho",
                "microfono",
                "camara",
                "reunion",
                "reuniones"
            ],
            respuesta:
                "SOPORTE DE MICROSOFT TEAMS\n\n" +
                "1. Comprueba que tu conexión a internet funcione.\n" +
                "2. Revisa los dispositivos de audio y cámara seleccionados.\n" +
                "3. Comprueba que Teams tenga los permisos necesarios para utilizar el micrófono.\n" +
                "4. Si es una reunión, verifica que no estés silenciado.\n\n" +
                "Si continúa el problema, indica si ocurre en todas las reuniones " +
                "o solo en una."
        },
        {
            id: "seguridad",
            categoria: "Seguridad",
            palabras: [
                "phishing",
                "hackearon",
                "hackeo",
                "correo sospechoso",
                "enlace sospechoso",
                "virus",
                "malware",
                "robaron mi cuenta",
                "acceso no autorizado",
                "me hackearon",
                "incidente de seguridad"
            ],
            respuesta:
                "POSIBLE INCIDENTE DE SEGURIDAD\n\n" +
                "1. No abras más enlaces ni archivos sospechosos.\n" +
                "2. No introduzcas contraseñas ni códigos en páginas dudosas.\n" +
                "3. Informa inmediatamente al área responsable de seguridad o TI.\n" +
                "4. Si ingresaste credenciales corporativas, comunícalo expresamente " +
                "para que se evalúe la cuenta y sus sesiones.\n\n" +
                "No borres evidencias del incidente ni intentes investigar al atacante " +
                "por tu cuenta."
        },
        {
            id: "inventario",
            categoria: "Infraestructura",
            palabras: [
                "inventario",
                "equipos conectados",
                "equipos desconectados",
                "monitoreo",
                "infraestructura",
                "estado de equipos",
                "computadoras registradas",
                "equipo asignado"
            ],
            respuesta:
                "CONSULTA DE INFRAESTRUCTURA\n\n" +
                "Puedo orientarte sobre consultas de inventario y monitoreo.\n\n" +
                "La conexión con las listas de equipos de ALFERZA aún no está habilitada " +
                "en esta versión. Por tanto, no puedo confirmar cantidades, estados " +
                "ni asignaciones reales.\n\n" +
                "Cuando integremos Microsoft Lists, podremos consultar la información " +
                "autorizada directamente."
        },
        {
            id: "gracias",
            categoria: "General",
            palabras: [
                "gracias",
                "muchas gracias",
                "te agradezco",
                "perfecto",
                "excelente",
                "entendido"
            ],
            respuesta:
                "¡Con gusto! Si necesitas ayuda con otra incidencia de TI, " +
                "descríbeme qué sucede y te orientaré."
        },
        {
            id: "despedida",
            categoria: "General",
            palabras: [
                "adios",
                "hasta luego",
                "nos vemos",
                "eso es todo",
                "nada mas"
            ],
            respuesta:
                "De acuerdo. Aquí estaré cuando necesites orientación de soporte TI."
        }
    ];

    const RESPUESTA_DESCONOCIDA =
        "Todavía no tengo un procedimiento específico para esa consulta.\n\n" +
        "Prueba describiendo el problema de otra manera o elige una categoría:\n\n" +
        "• Instalación de software.\n" +
        "• Conexión a internet.\n" +
        "• Outlook y correo.\n" +
        "• Impresoras y equipos.\n" +
        "• Solicitudes de soporte.\n\n" +
        "Esta versión utiliza reglas y respuestas predefinidas; aún no dispone " +
        "de una red neuronal entrenada ni consulta datos en tiempo real.";

    /* =====================================================
       2. NORMALIZACIÓN DEL TEXTO
    ===================================================== */

    function normalizarTexto(texto) {
        return String(texto || "")
            .toLocaleLowerCase("es")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[^\p{L}\p{N}\s]/gu, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    /* =====================================================
       3. DETECCIÓN DE INTENCIONES
    ===================================================== */

    function contieneFrase(texto, frase) {
        const consulta = ` ${texto} `;
        const palabra = ` ${normalizarTexto(frase)} `;

        return consulta.includes(palabra);
    }

    function calcularCoincidencia(consulta, intencion) {
        let puntuacion = 0;

        for (const palabra of intencion.palabras) {
            const termino = normalizarTexto(palabra);

            if (!termino) continue;

            if (contieneFrase(consulta, termino)) {
                // Prioriza expresiones específicas sobre palabras genéricas.
                puntuacion += termino.includes(" ")
                    ? 3 + termino.split(" ").length
                    : 1;
            }
        }

        return puntuacion;
    }

    function detectarIntencion(mensaje) {
        const consulta = normalizarTexto(mensaje);

        if (!consulta) {
            return null;
        }

        let mejorIntencion = null;
        let mejorPuntuacion = 0;

        for (const intencion of BASE_CONOCIMIENTO) {
            const puntuacion = calcularCoincidencia(
                consulta,
                intencion
            );

            if (puntuacion > mejorPuntuacion) {
                mejorPuntuacion = puntuacion;
                mejorIntencion = intencion;
            }
        }

        // Evita clasificar frases con coincidencias demasiado débiles.
        if (mejorPuntuacion < 1) {
            return null;
        }

        return mejorIntencion;
    }

    function generarRespuesta(mensaje) {
        const intencion = detectarIntencion(mensaje);

        if (!intencion) {
            return {
                texto: RESPUESTA_DESCONOCIDA,
                categoria: "Sin clasificar",
                intencion: "desconocida"
            };
        }

        return {
            texto: intencion.respuesta,
            categoria: intencion.categoria,
            intencion: intencion.id
        };
    }

    /* =====================================================
       4. RENDERIZADO SEGURO DE MENSAJES
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

        // textContent evita interpretar mensajes como HTML ejecutable.
        burbuja.textContent = texto;

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
        elements.messages.scrollTop = elements.messages.scrollHeight;
    }

    /* =====================================================
       5. ESTADO DE LA INTERFAZ
    ===================================================== */

    function mostrarEscritura(visible) {
        elements.typing.hidden = !visible;
        elements.sendButton.disabled = visible;
        elements.input.disabled = visible;

        if (visible) {
            desplazarAlFinal();
        }
    }

    function actualizarContador() {
        const cantidad = elements.input.value.length;

        elements.characterCount.textContent = `${cantidad}/1000`;

        // Ajusta la altura del cuadro de texto automáticamente.
        elements.input.style.height = "auto";
        elements.input.style.height =
            `${Math.min(elements.input.scrollHeight, 120)}px`;
    }

    /* =====================================================
       6. PROCESAMIENTO DE MENSAJES
    ===================================================== */

    function enviarMensaje(mensaje) {
        const texto = String(mensaje || "").trim();

        if (!texto || procesandoMensaje) {
            return;
        }

        procesandoMensaje = true;

        agregarMensaje("user", texto);

        elements.input.value = "";
        actualizarContador();
        mostrarEscritura(true);

        // Simula un breve tiempo de respuesta para la interfaz.
        // No es una llamada a una IA ni a un servidor.
        window.setTimeout(function () {
            const resultado = generarRespuesta(texto);

            mostrarEscritura(false);
            agregarMensaje("bot", resultado.texto);

            console.info("[ALFERZA CHATBOT]", {
                intencion: resultado.intencion,
                categoria: resultado.categoria
            });

            procesandoMensaje = false;
            elements.input.disabled = false;
            elements.sendButton.disabled = false;
            elements.input.focus();
        }, 350);
    }

    /* =====================================================
       7. EVENTOS
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

    elements.clearButton.addEventListener("click", function () {
        if (procesandoMensaje) {
            return;
        }

        elements.messages.replaceChildren();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "¿Qué problema necesitas resolver hoy?"
        );

        elements.input.value = "";
        actualizarContador();
        elements.input.focus();
    });

    /* =====================================================
       8. INICIALIZACIÓN
    ===================================================== */

    function iniciarChatbot() {
        actualizarContador();

        agregarMensaje(
            "bot",
            "¡Hola! Soy el asistente de TI de ALFERZA.\n\n" +
            "Puedo orientarte con instalaciones, problemas de red, " +
            "Outlook, impresoras y solicitudes de soporte.\n\n" +
            "Cuéntame qué necesitas resolver."
        );

        console.info(
            "[ALFERZA CHATBOT] Inicializado correctamente."
        );
    }

    iniciarChatbot();
})();
