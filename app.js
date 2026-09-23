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
   ELEMENTOS DE CARGA
   ========================================= */

const overlay =
    document.getElementById("loadingOverlay");

const loadingMessage =
    document.getElementById("loadingMessage");


function cambiarMensaje(mensaje) {

    if (loadingMessage) {

        loadingMessage.textContent =
            mensaje;

    }

}


/* =========================================
   MOSTRAR / OCULTAR CARGA
   ========================================= */

function mostrarCarga(mensaje) {

    cambiarMensaje(mensaje);

    overlay.classList.remove("oculto");

}


function ocultarCarga() {

    overlay.classList.add("oculto");

}


/* =========================================
   OBTENER TOKEN
   ========================================= */

async function obtenerToken() {

    let cuenta =
        msalInstance.getAllAccounts()[0];


    /* -----------------------------------------
       NO HAY SESIÓN
       ----------------------------------------- */

    if (!cuenta) {

        cambiarMensaje(
            "Autenticando con Microsoft..."
        );

        const loginResponse =
            await msalInstance.loginPopup({
                scopes
            });

        cuenta =
            loginResponse.account;

    }


    /* -----------------------------------------
       INTENTAR TOKEN SILENCIOSO
       ----------------------------------------- */

    try {

        cambiarMensaje(
            "Verificando permisos..."
        );

        const response =
            await msalInstance.acquireTokenSilent({

                scopes,

                account: cuenta

            });

        return response.accessToken;

    }


    /* -----------------------------------------
       TOKEN SILENCIOSO FALLA
       ----------------------------------------- */

    catch (error) {

        cambiarMensaje(
            "Renovando sesión..."
        );

        const response =
            await msalInstance.acquireTokenPopup({
                scopes
            });

        return response.accessToken;

    }

}


/* =========================================
   CARGAR USUARIOS
   ========================================= */

async function cargarUsuarios() {

    mostrarCarga(
        "Cargando información del personal..."
    );


    const contenedor =
        document.getElementById("officeGrid");


    contenedor.innerHTML = "";


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


        const respuesta =
            await fetch(
                "https://graph.microsoft.com/v1.0/users?$top=999",
                {
                    headers: {
                        Authorization:
                            `Bearer ${TOKEN}`
                    }
                }
            );


        if (!respuesta.ok) {

            throw new Error(
                `Error Microsoft Graph: ${respuesta.status}`
            );

        }


        const data =
            await respuesta.json();


        const usuarios =
            data.value.filter(
                u =>
                    u.mail &&
                    u.mail
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
                u => u.id
            );


        let presenciaPorId = {};


        /*
         * Microsoft Graph permite hasta
         * 650 IDs por solicitud.
         *
         * Si en el futuro tienes más de 650,
         * se pueden dividir en bloques.
         */

        const bloques = [];

        for (
            let i = 0;
            i < idsUsuarios.length;
            i += 650
        ) {

            bloques.push(
                idsUsuarios.slice(
                    i,
                    i + 650
                )
            );

        }


        for (const bloque of bloques) {

            try {

                const presenciaResponse =
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


                if (!presenciaResponse.ok) {

                    continue;

                }


                const presenciaData =
                    await presenciaResponse.json();


                (
                    presenciaData.value || []
                ).forEach(p => {

                    presenciaPorId[p.id] = p;

                });

            }

            catch (error) {

                console.warn(
                    "No se pudo obtener presencia:",
                    error
                );

            }

        }


        /* -----------------------------------------
           GENERAR TARJETAS
           ----------------------------------------- */

        const usuariosConPresencia =
            usuarios.map(usuario => ({

                usuario,

                presencia:
                    presenciaPorId[
                        usuario.id
                    ] || {
                        availability:
                            "Offline"
                    }

            }));


        let disponibles = 0;

        let ocupados = 0;

        let ausentes = 0;

        let offline = 0;


        let html = "";


        usuariosConPresencia.forEach(item => {

            const usuario =
                item.usuario;


            const presencia =
                item.presencia;


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


        /* -----------------------------------------
           ESTADÍSTICAS
           ----------------------------------------- */

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


        /* -----------------------------------------
           FINALIZAR CARGA
           ----------------------------------------- */

        cambiarMensaje(
            "Información actualizada"
        );


        /*
         * Pequeña pausa para que no desaparezca
         * bruscamente la pantalla.
         */

        setTimeout(() => {

            ocultarCarga();

        }, 250);


    }


    catch (error) {

        console.error(
            "Error cargando ALFERZA LIVE OFFICE:",
            error
        );


        cambiarMensaje(
            "No se pudo cargar la información"
        );


        /*
         * Mantener la pantalla de seguridad.
         * No mostramos información parcial.
         */

        const contenedor =
            document.getElementById("officeGrid");


        contenedor.innerHTML = "";


        /*
         * Después de unos segundos mostramos
         * un mensaje de error.
         */

        setTimeout(() => {

            loadingMessage.innerHTML =
                "No se pudo conectar con Microsoft.<br>" +
                "<small>Recarga la página para intentarlo nuevamente.</small>";

        }, 500);

    }

}


/* =========================================
   INICIAR APLICACIÓN
   ========================================= */

cargarUsuarios();


/* =========================================
   ACTUALIZACIÓN AUTOMÁTICA
   ========================================= */

setInterval(() => {

    cargarUsuarios();

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