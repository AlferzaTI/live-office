/* =========================================
   CONFIGURACIÓN MSAL
   ========================================= */

const msalConfig = {

    auth: {
        clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",

        authority:
            "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",

        redirectUri:
            "https://alvaroalferza.github.io/live-office/blank.html"
    },

    cache: {
        cacheLocation: "sessionStorage"
    }

};


const scopes = [
    "User.Read",
    "Presence.Read.All",
    "Sites.Read.All"
];


const msalInstance =
    new msal.PublicClientApplication(msalConfig);


/* =========================================
   ELEMENTOS
   ========================================= */

const overlay =
    document.getElementById("loadingOverlay");

const loadingMessage =
    document.getElementById("loadingMessage");


/* =========================================
   CONFIGURACIÓN DE SESIÓN
   ========================================= */

/*
 * IMPORTANTE:
 *
 * Usamos localStorage y NO sessionStorage.
 *
 * De esta manera la validación se conserva:
 *
 * - al actualizar Index
 * - al salir a Personal
 * - al volver a Index
 * - al entrar a Reservas
 * - al volver a Index
 * - al cerrar y abrir nuevamente el navegador
 *
 * Mientras no borremos esta marca,
 * no vuelve a mostrar la validación.
 */

const VALIDACION_KEY =
    "alferza_live_office_validado";


function estaValidado() {

    return localStorage.getItem(
        VALIDACION_KEY
    ) === "true";

}


function guardarValidacion() {

    localStorage.setItem(
        VALIDACION_KEY,
        "true"
    );

}


/* =========================================
   MENSAJE DE CARGA
   ========================================= */

function cambiarMensaje(mensaje) {

    if (loadingMessage) {

        loadingMessage.textContent =
            mensaje;

    }

}


/* =========================================
   MOSTRAR CARGA
   ========================================= */

function mostrarCarga(mensaje) {

    if (!overlay) {
        return;
    }

    cambiarMensaje(mensaje);

    overlay.classList.remove("oculto");

}


/* =========================================
   OCULTAR CARGA
   ========================================= */

function ocultarCarga() {

    if (!overlay) {
        return;
    }

    overlay.classList.add("oculto");

}


/* =========================================
   ESPERAR
   ========================================= */

function esperar(ms) {

    return new Promise(resolve => {

        setTimeout(resolve, ms);

    });

}


/* =========================================
   FETCH CON TIEMPO MÁXIMO
   ========================================= */

/*
 * Evita que Microsoft Graph deje la página
 * cargando eternamente.
 */

async function fetchConTimeout(
    url,
    opciones = {},
    tiempo = 20000
) {

    const controller =
        new AbortController();


    const timeout =
        setTimeout(() => {

            controller.abort();

        }, tiempo);


    try {

        const respuesta =
            await fetch(
                url,
                {
                    ...opciones,
                    signal:
                        controller.signal
                }
            );


        return respuesta;

    }


    catch (error) {

        if (
            error.name ===
            "AbortError"
        ) {

            throw new Error(
                "La conexión con Microsoft tardó demasiado."
            );

        }


        throw error;

    }


    finally {

        clearTimeout(timeout);

    }

}


/* =========================================
   OBTENER TOKEN
   ========================================= */

async function obtenerToken() {

    let cuenta =
        msalInstance.getAllAccounts()[0];


    /* -----------------------------------------
       SI NO HAY CUENTA
       ----------------------------------------- */

    if (!cuenta) {

        cambiarMensaje(
            "Autenticando con Microsoft..."
        );


        /*
         * Primera vez solamente.
         */

        const loginResponse =
            await Promise.race([

                msalInstance.loginPopup({
                    scopes: scopes
                }),

                new Promise(
                    (_, reject) => {

                        setTimeout(() => {

                            reject(
                                new Error(
                                    "La autenticación de Microsoft tardó demasiado."
                                )
                            );

                        }, 60000);

                    }
                )

            ]);


        cuenta =
            loginResponse.account;


        /*
         * Guardamos la cuenta activa.
         */

        msalInstance.setActiveAccount(
            cuenta
        );

    }


    else {

        /*
         * Nos aseguramos de tener
         * una cuenta activa.
         */

        msalInstance.setActiveAccount(
            cuenta
        );

    }


    /* -----------------------------------------
       TOKEN SILENCIOSO
       ----------------------------------------- */

    try {

        cambiarMensaje(
            "Verificando permisos..."
        );


        const respuesta =
            await Promise.race([

                msalInstance.acquireTokenSilent({

                    scopes: scopes,

                    account: cuenta

                }),

                new Promise(
                    (_, reject) => {

                        setTimeout(() => {

                            reject(
                                new Error(
                                    "La validación de Microsoft tardó demasiado."
                                )
                            );

                        }, 30000);

                    }
                )

            ]);


        return respuesta.accessToken;

    }


    catch (error) {

        console.warn(
            "No se pudo obtener el token silenciosamente:",
            error
        );


        /*
         * Solo intentamos popup si realmente
         * Microsoft necesita interacción.
         */

        cambiarMensaje(
            "Actualizando sesión de Microsoft..."
        );


        const respuesta =
            await Promise.race([

                msalInstance.acquireTokenPopup({
                    scopes: scopes
                }),

                new Promise(
                    (_, reject) => {

                        setTimeout(() => {

                            reject(
                                new Error(
                                    "Microsoft no respondió a tiempo."
                                )
                            );

                        }, 60000);

                    }
                )

            ]);


        return respuesta.accessToken;

    }

}


/* =========================================
   OBTENER USUARIOS
   ========================================= */

async function obtenerUsuarios(TOKEN) {

    const respuesta =
        await fetchConTimeout(

            "https://graph.microsoft.com/v1.0/users?$top=999",

            {
                method: "GET",

                headers: {

                    Authorization:
                        `Bearer ${TOKEN}`

                }

            },

            20000

        );


    if (!respuesta.ok) {

        throw new Error(
            `Microsoft Graph Users: HTTP ${respuesta.status}`
        );

    }


    return await respuesta.json();

}


/* =========================================
   OBTENER PRESENCIA
   ========================================= */

async function obtenerPresencia(
    TOKEN,
    idsUsuarios
) {

    const presenciaPorId = {};


    /*
     * Microsoft Graph permite hasta
     * 650 usuarios por solicitud.
     */

    for (
        let i = 0;
        i < idsUsuarios.length;
        i += 650
    ) {

        const bloque =
            idsUsuarios.slice(
                i,
                i + 650
            );


        const respuesta =
            await fetchConTimeout(

                "https://graph.microsoft.com/v1.0/communications/getPresencesByUserId",

                {
                    method: "POST",

                    headers: {

                        Authorization:
                            `Bearer ${TOKEN}`,

                        "Content-Type":
                            "application/json"

                    },

                    body:
                        JSON.stringify({
                            ids: bloque
                        })

                },

                20000

            );


        if (!respuesta.ok) {

            throw new Error(
                `Microsoft Graph Presence: HTTP ${respuesta.status}`
            );

        }


        const data =
            await respuesta.json();


        (
            data.value || []
        ).forEach(presencia => {

            presenciaPorId[
                presencia.id
            ] = presencia;

        });

    }


    return presenciaPorId;

}


/* =========================================
   RENDERIZAR PERSONAL
   ========================================= */

function renderizarUsuarios(
    usuarios,
    presenciaPorId
) {

    const contenedor =
        document.getElementById(
            "officeGrid"
        );


    if (!contenedor) {
        return;
    }


    let disponibles = 0;
    let ocupados = 0;
    let ausentes = 0;
    let offline = 0;


    let html = "";


    usuarios.forEach(usuario => {

        const presencia =
            presenciaPorId[
                usuario.id
            ] || {
                availability: "Offline"
            };


        const estado =
            presencia.availability ||
            "Offline";


        let clase =
            "offline";


        switch (estado) {

            case "Available":

                clase =
                    "disponible";

                disponibles++;

                break;


            case "Busy":

            case "InAMeeting":

            case "OnACall":

                clase =
                    "ocupado";

                ocupados++;

                break;


            case "Away":

            case "BeRightBack":

                clase =
                    "ausente";

                ausentes++;

                break;


            default:

                clase =
                    "offline";

                offline++;

                break;

        }


        html += `

            <div
                class="card"
                data-estado="${estado}"
            >

                <h3>
                    ${usuario.displayName}
                </h3>

                <p>
                    ${usuario.mail}
                </p>

                <div
                    class="status ${clase}"
                >
                    ${estado}
                </div>

            </div>

        `;

    });


    contenedor.innerHTML =
        html;


    const disp =
        document.getElementById("disp");

    const busy =
        document.getElementById("busy");

    const away =
        document.getElementById("away");

    const offlineElement =
        document.getElementById("offline");


    if (disp) {

        disp.innerText =
            disponibles;

    }


    if (busy) {

        busy.innerText =
            ocupados;

    }


    if (away) {

        away.innerText =
            ausentes;

    }


    if (offlineElement) {

        offlineElement.innerText =
            offline;

    }

}


/* =========================================
   CARGAR INFORMACIÓN
   ========================================= */

async function cargarUsuarios(
    mostrarPantalla = false
) {

    /*
     * La pantalla azul SOLO se muestra
     * si es la primera validación.
     */

    if (mostrarPantalla) {

        mostrarCarga(
            "Verificando acceso..."
        );

    }


    try {

        /* -----------------------------------------
           AUTENTICACIÓN
           ----------------------------------------- */

        const TOKEN =
            await obtenerToken();


        /* -----------------------------------------
           USUARIOS
           ----------------------------------------- */

        cambiarMensaje(
            "Consultando personal..."
        );


        const data =
            await obtenerUsuarios(
                TOKEN
            );


        const usuarios =
            (data.value || []).filter(
                usuario =>

                    usuario.mail &&

                    usuario.mail
                        .toLowerCase()
                        .endsWith("@alferza.pe")
            );


        /* -----------------------------------------
           PRESENCIA
           ----------------------------------------- */

        cambiarMensaje(
            "Consultando estados..."
        );


        const idsUsuarios =
            usuarios.map(
                usuario =>
                    usuario.id
            );


        const presenciaPorId =
            await obtenerPresencia(
                TOKEN,
                idsUsuarios
            );


        /* -----------------------------------------
           RENDERIZAR
           ----------------------------------------- */

        renderizarUsuarios(
            usuarios,
            presenciaPorId
        );


        /*
         * TODO salió correctamente.
         *
         * Desde este momento ya no se
         * vuelve a mostrar la pantalla azul.
         */

        guardarValidacion();


        /* -----------------------------------------
           OCULTAR OVERLAY
           ----------------------------------------- */

        if (mostrarPantalla) {

            cambiarMensaje(
                "Información actualizada"
            );


            await esperar(200);


            ocultarCarga();

        }


        console.log(
            "ALFERZA LIVE OFFICE actualizado correctamente."
        );

    }


    catch (error) {

        console.error(
            "Error cargando ALFERZA LIVE OFFICE:",
            error
        );


        /*
         * Si era una actualización normal,
         * NO bloqueamos la pantalla.
         */

        if (!mostrarPantalla) {

            console.warn(
                "Actualización silenciosa fallida."
            );

            return;

        }


        /*
         * Si era la primera validación,
         * mostramos el error.
         */

        cambiarMensaje(
            "No se pudo conectar con Microsoft"
        );


        /*
         * Evita que el usuario quede
         * atrapado eternamente en la pantalla.
         */

        setTimeout(() => {

            ocultarCarga();

        }, 2500);

    }

}


/* =========================================
   INICIO
   ========================================= */

/*
 * PRIMERA VEZ:
 *
 *   mostrarPantalla = true
 *
 * Después de validar:
 *
 *   localStorage = true
 *
 *
 * SIGUIENTES VECES:
 *
 *   mostrarPantalla = false
 *
 *
 * Por lo tanto:
 *
 * Index
 * ↓
 * Personal
 * ↓
 * Reservas
 * ↓
 * Salas
 * ↓
 * Comunicados
 * ↓
 * Index
 *
 * NO vuelve a aparecer la validación.
 *
 *
 * Incluso si presionas:
 *
 * Ctrl + R
 *
 * tampoco vuelve a aparecer.
 */

if (estaValidado()) {

    cargarUsuarios(false);

}
else {

    cargarUsuarios(true);

}


/* =========================================
   ACTUALIZACIÓN CADA 5 MINUTOS
   ========================================= */

setInterval(() => {

    /*
     * Siempre silencioso.
     *
     * No aparece la pantalla azul.
     */

    cargarUsuarios(false);

}, 300000);


/* =========================================
   BUSCADOR
   ========================================= */

const buscador =
    document.getElementById(
        "buscador"
    );


if (buscador) {

    buscador.addEventListener(
        "keyup",
        function () {

            const texto =
                this.value.toLowerCase();


            document
                .querySelectorAll(".card")
                .forEach(card => {

                    const contenido =
                        card.innerText
                            .toLowerCase();


                    card.style.display =
                        contenido.includes(texto)
                            ? ""
                            : "none";

                });

        }
    );

}


/* =========================================
   FILTROS
   ========================================= */

function filtrarEstado(tipo) {

    document
        .querySelectorAll(".card")
        .forEach(card => {

            const estado =
                card.dataset.estado;


            let mostrar =
                false;


            switch (tipo) {

                case "Available":

                    mostrar =
                        estado === "Available";

                    break;


                case "Busy":

                    mostrar =
                        estado === "Busy" ||
                        estado === "InAMeeting" ||
                        estado === "OnACall";

                    break;


                case "Away":

                    mostrar =
                        estado === "Away" ||
                        estado === "BeRightBack";

                    break;


                case "Offline":

                    mostrar =
                        estado === "Offline";

                    break;


                default:

                    mostrar =
                        true;

            }


            card.style.display =
                mostrar
                    ? ""
                    : "none";

        });

}


/* =========================================
   MOSTRAR TODOS
   ========================================= */

function mostrarTodos() {

    document
        .querySelectorAll(".card")
        .forEach(card => {

            card.style.display =
                "";

        });

}