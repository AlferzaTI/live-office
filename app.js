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

    cambiarMensaje(mensaje);

    overlay.classList.remove("oculto");

}


/* =========================================
   OCULTAR CARGA
   ========================================= */

function ocultarCarga() {

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
   OBTENER TOKEN
   ========================================= */

async function obtenerToken() {

    let cuenta =
        msalInstance.getAllAccounts()[0];


    /* -----------------------------------------
       SI NO EXISTE SESIÓN
       ----------------------------------------- */

    if (!cuenta) {

        cambiarMensaje(
            "Autenticando con Microsoft..."
        );


        try {

            const loginResponse =
                await msalInstance.loginPopup({
                    scopes: scopes
                });


            cuenta =
                loginResponse.account;

        }

        catch (error) {

            console.error(
                "Error durante el inicio de sesión:",
                error
            );

            throw error;

        }

    }


    /* -----------------------------------------
       TOKEN SILENCIOSO
       ----------------------------------------- */

    try {

        cambiarMensaje(
            "Verificando permisos..."
        );


        const response =
            await msalInstance.acquireTokenSilent({

                scopes: scopes,

                account: cuenta

            });


        return response.accessToken;

    }


    catch (error) {

        console.warn(
            "Token silencioso no disponible:",
            error
        );


        /*
         * Solo en caso de que Microsoft
         * necesite interacción nuevamente.
         */

        cambiarMensaje(
            "Actualizando sesión..."
        );


        const response =
            await msalInstance.acquireTokenPopup({
                scopes: scopes
            });


        return response.accessToken;

    }

}


/* =========================================
   OBTENER USUARIOS DESDE GRAPH
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

    let presenciaPorId = {};


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


    document.getElementById("disp")
        .innerText =
        disponibles;


    document.getElementById("busy")
        .innerText =
        ocupados;


    document.getElementById("away")
        .innerText =
        ausentes;


    document.getElementById("offline")
        .innerText =
        offline;

}


/* =========================================
   CARGAR INFORMACIÓN
   ========================================= */

async function cargarUsuarios(
    mostrarPantalla = true
) {

    /*
     * Solo mostramos el overlay durante
     * la carga inicial.
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
            data.value.filter(
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
           MOSTRAR INFORMACIÓN
           ----------------------------------------- */

        renderizarUsuarios(
            usuarios,
            presenciaPorId
        );


        /* -----------------------------------------
           OCULTAR PANTALLA
           ----------------------------------------- */

        if (mostrarPantalla) {

            cambiarMensaje(
                "Información actualizada"
            );


            await esperar(250);


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
         * Si es la carga inicial,
         * NO mostramos información parcial.
         */

        if (mostrarPantalla) {

            cambiarMensaje(
                "Conectando con Microsoft..."
            );


            /*
             * Reintentar automáticamente.
             *
             * Esto evita que el usuario tenga
             * que hacer Ctrl + Shift + R.
             */

            await esperar(1500);


            try {

                const TOKEN =
                    await obtenerToken();


                const data =
                    await obtenerUsuarios(
                        TOKEN
                    );


                const usuarios =
                    data.value.filter(
                        usuario =>
                            usuario.mail &&
                            usuario.mail
                                .toLowerCase()
                                .endsWith(
                                    "@alferza.pe"
                                )
                    );


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


                renderizarUsuarios(
                    usuarios,
                    presenciaPorId
                );


                cambiarMensaje(
                    "Información actualizada"
                );


                await esperar(250);


                ocultarCarga();


                console.log(
                    "Conexión recuperada automáticamente."
                );


                return;

            }


            catch (errorSegundoIntento) {

                console.error(
                    "Segundo intento fallido:",
                    errorSegundoIntento
                );


                cambiarMensaje(
                    "No se pudo conectar con Microsoft"
                );


                return;

            }

        }

    }

}


/* =========================================
   INICIO
   ========================================= */

cargarUsuarios(true);


/* =========================================
   ACTUALIZACIÓN CADA 5 MINUTOS
   ========================================= */

setInterval(() => {

    /*
     * IMPORTANTE:
     * false = NO mostrar pantalla azul.
     *
     * Los usuarios siguen viendo
     * la aplicación mientras se actualiza.
     */

    cargarUsuarios(false);

}, 300000);


/* =========================================
   BUSCADOR
   ========================================= */

document
    .getElementById("buscador")
    .addEventListener(
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