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
   CONTROL DE SESIÓN
   ========================================= */

/*
 * Esta marca se guarda en sessionStorage.
 *
 * Mientras la pestaña siga abierta:
 *
 * Index → Personal → Reservas → Salas → Index
 *
 * NO vuelve a aparecer la pantalla azul.
 */

const SESION_KEY =
    "alferza_live_office_validado";


function sesionValidada() {

    return sessionStorage.getItem(
        SESION_KEY
    ) === "true";

}


function guardarSesionValidada() {

    sessionStorage.setItem(
        SESION_KEY,
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
   OBTENER TOKEN
   ========================================= */

async function obtenerToken() {

    let cuenta =
        msalInstance.getAllAccounts()[0];


    /*
     * Si no existe una cuenta,
     * abrimos Microsoft UNA SOLA VEZ.
     */

    if (!cuenta) {

        cambiarMensaje(
            "Autenticando con Microsoft..."
        );


        const loginResponse =
            await msalInstance.loginPopup({
                scopes: scopes
            });


        cuenta =
            loginResponse.account;

    }


    /*
     * Intentamos obtener el token
     * silenciosamente.
     */

    try {

        cambiarMensaje(
            "Verificando permisos..."
        );


        const respuesta =
            await msalInstance.acquireTokenSilent({

                scopes: scopes,

                account: cuenta

            });


        return respuesta.accessToken;

    }


    catch (error) {

        console.warn(
            "Token silencioso no disponible. Se solicitará nuevamente.",
            error
        );


        cambiarMensaje(
            "Actualizando sesión..."
        );


        const respuesta =
            await msalInstance.acquireTokenPopup({

                scopes: scopes

            });


        return respuesta.accessToken;

    }

}


/* =========================================
   OBTENER USUARIOS
   ========================================= */

async function obtenerUsuarios(TOKEN) {

    const respuesta =
        await fetch(
            "https://graph.microsoft.com/v1.0/users?$top=999",
            {
                method: "GET",

                headers: {
                    Authorization:
                        `Bearer ${TOKEN}`
                }
            }
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
     * 650 IDs por solicitud.
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
            await fetch(
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

                }
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
        document.getElementById("officeGrid");


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
     * SOLO mostramos la pantalla azul
     * si realmente es la primera entrada.
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
         * Guardamos la sesión.
         */

        guardarSesionValidada();


        /* -----------------------------------------
           OCULTAR OVERLAY
           ----------------------------------------- */

        if (mostrarPantalla) {

            cambiarMensaje(
                "Información actualizada"
            );


            /*
             * Una pausa MUY corta para
             * que la transición sea visible.
             */

            setTimeout(() => {

                ocultarCarga();

            }, 200);

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
         * Si la carga NO era inicial,
         * no mostramos ninguna pantalla.
         *
         * La aplicación continúa funcionando.
         */

        if (!mostrarPantalla) {

            return;

        }


        /*
         * Si era la primera entrada,
         * mostramos el error.
         */

        cambiarMensaje(
            "No se pudo conectar con Microsoft"
        );


        /*
         * Quitamos el overlay después de unos
         * segundos para evitar que quede
         * bloqueado eternamente.
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
 * AQUÍ ESTÁ EL CAMBIO PRINCIPAL.
 *
 * Si la sesión ya fue validada:
 *
 *     cargarUsuarios(false)
 *
 * No aparece la pantalla azul.
 *
 *
 * Si es la primera entrada:
 *
 *     cargarUsuarios(true)
 *
 * Aparece la pantalla azul.
 */

if (sesionValidada()) {

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
     * No aparece pantalla azul.
     */

    cargarUsuarios(false);

}, 300000);


/* =========================================
   BUSCADOR
   ========================================= */

const buscador =
    document.getElementById("buscador");


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