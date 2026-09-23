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
   CLAVES DE ALMACENAMIENTO
   ========================================= */

/*
 * VALIDACION:
 *
 * Indica que el usuario ya pasó por Microsoft.
 *
 * DATOS:
 *
 * Guarda la última información obtenida
 * para que Index pueda mostrarse inmediatamente
 * cuando el usuario vuelva.
 */

const VALIDACION_KEY =
    "alferza_live_office_validado";

const DATOS_KEY =
    "alferza_live_office_datos";


/* =========================================
   COMPROBAR SI YA FUE VALIDADO
   ========================================= */

function yaEstaValidado() {

    return localStorage.getItem(
        VALIDACION_KEY
    ) === "true";

}


/* =========================================
   GUARDAR VALIDACIÓN
   ========================================= */

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
   GUARDAR DATOS
   ========================================= */

function guardarDatos(usuarios, presenciaPorId) {

    try {

        localStorage.setItem(

            DATOS_KEY,

            JSON.stringify({

                usuarios:
                    usuarios,

                presencia:
                    presenciaPorId

            })

        );

    }

    catch (error) {

        console.warn(
            "No se pudieron guardar los datos:",
            error
        );

    }

}


/* =========================================
   RECUPERAR DATOS
   ========================================= */

function recuperarDatos() {

    try {

        const datos =
            localStorage.getItem(
                DATOS_KEY
            );


        if (!datos) {

            return null;

        }


        return JSON.parse(datos);

    }

    catch (error) {

        console.warn(
            "No se pudieron recuperar los datos:",
            error
        );

        return null;

    }

}


/* =========================================
   OBTENER TOKEN
   ========================================= */

async function obtenerToken() {

    let cuenta =
        msalInstance.getAllAccounts()[0];


    /* -----------------------------------------
       PRIMERA VEZ
       ----------------------------------------- */

    if (!cuenta) {

        cambiarMensaje(
            "Autenticando con Microsoft..."
        );


        const loginResponse =
            await msalInstance.loginPopup({

                scopes:
                    scopes

            });


        cuenta =
            loginResponse.account;

    }


    /* -----------------------------------------
       CUENTA ACTIVA
       ----------------------------------------- */

    msalInstance.setActiveAccount(
        cuenta
    );


    /* -----------------------------------------
       TOKEN SILENCIOSO
       ----------------------------------------- */

    try {

        const respuesta =
            await msalInstance.acquireTokenSilent({

                scopes:
                    scopes,

                account:
                    cuenta

            });


        return respuesta.accessToken;

    }


    catch (error) {

        console.warn(
            "Token silencioso no disponible:",
            error
        );


        /*
         * Solo Microsoft decide si necesita
         * volver a pedir interacción.
         */

        const respuesta =
            await msalInstance.acquireTokenPopup({

                scopes:
                    scopes

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

                method:
                    "GET",

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
     * Hasta 650 usuarios por petición.
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

                    method:
                        "POST",

                    headers: {

                        Authorization:
                            `Bearer ${TOKEN}`,

                        "Content-Type":
                            "application/json"

                    },

                    body:
                        JSON.stringify({

                            ids:
                                bloque

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
            ] =
                presencia;

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


    let disponibles =
        0;

    let ocupados =
        0;

    let ausentes =
        0;

    let offline =
        0;


    let html =
        "";


    usuarios.forEach(usuario => {

        const presencia =
            presenciaPorId[
                usuario.id
            ] || {

                availability:
                    "Offline"

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
        document.getElementById(
            "disp"
        );


    const busy =
        document.getElementById(
            "busy"
        );


    const away =
        document.getElementById(
            "away"
        );


    const offlineElement =
        document.getElementById(
            "offline"
        );


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
   CARGAR DATOS DE INTERNET
   ========================================= */

async function actualizarDesdeMicrosoft(
    mostrarPantalla
) {

    try {

        /* -----------------------------------------
           SOLO PRIMERA VEZ
           ----------------------------------------- */

        if (mostrarPantalla) {

            mostrarCarga(
                "Verificando acceso..."
            );

        }


        /* -----------------------------------------
           TOKEN
           ----------------------------------------- */

        const TOKEN =
            await obtenerToken();


        /* -----------------------------------------
           USUARIOS
           ----------------------------------------- */

        if (mostrarPantalla) {

            cambiarMensaje(
                "Consultando personal..."
            );

        }


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
                        .endsWith(
                            "@alferza.pe"
                        )

            );


        /* -----------------------------------------
           PRESENCIA
           ----------------------------------------- */

        if (mostrarPantalla) {

            cambiarMensaje(
                "Consultando estados..."
            );

        }


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


        /* -----------------------------------------
           GUARDAR DATOS
           ----------------------------------------- */

        guardarDatos(

            usuarios,

            presenciaPorId

        );


        /* -----------------------------------------
           MARCAR VALIDACIÓN
           ----------------------------------------- */

        guardarValidacion();


        /* -----------------------------------------
           OCULTAR PANTALLA
           ----------------------------------------- */

        if (mostrarPantalla) {

            cambiarMensaje(
                "Información actualizada"
            );


            ocultarCarga();

        }


        console.log(
            "Información de Microsoft actualizada."
        );


    }

    catch (error) {

        console.error(
            "Error actualizando información:",
            error
        );


        /*
         * SI ES LA PRIMERA VEZ:
         *
         * No podemos mostrar información
         * que todavía no tenemos.
         */

        if (mostrarPantalla) {

            cambiarMensaje(
                "No se pudo conectar con Microsoft"
            );


            /*
             * Quitamos la pantalla.
             * NO dejamos el Index bloqueado.
             */

            setTimeout(() => {

                ocultarCarga();

            }, 1500);

        }


        /*
         * SI NO ES LA PRIMERA VEZ:
         *
         * NO HACEMOS NADA.
         *
         * Los datos anteriores permanecen
         * visibles.
         */

    }

}


/* =========================================
   INICIO DEL INDEX
   ========================================= */

const datosGuardados =
    recuperarDatos();


/*
 * =========================================
 * CASO 1:
 * YA FUE VALIDADO ANTERIORMENTE
 * =========================================
 */

if (
    yaEstaValidado() &&
    datosGuardados
) {

    /*
     * MOSTRAR INMEDIATAMENTE
     *
     * No esperamos a Microsoft.
     * No mostramos overlay.
     * No bloqueamos Index.
     */

    renderizarUsuarios(

        datosGuardados.usuarios,

        datosGuardados.presencia

    );


    /*
     * Actualizamos Microsoft
     * EN SEGUNDO PLANO.
     *
     * El usuario ni siquiera verá
     * una pantalla de carga.
     */

    actualizarDesdeMicrosoft(
        false
    );

}


/*
 * =========================================
 * CASO 2:
 * YA FUE VALIDADO PERO NO HAY DATOS
 * =========================================
 */

else if (
    yaEstaValidado()
) {

    /*
     * No mostramos pantalla azul.
     *
     * Intentamos recuperar información
     * silenciosamente.
     */

    actualizarDesdeMicrosoft(
        false
    );

}


/*
 * =========================================
 * CASO 3:
 * PRIMERA VEZ
 * =========================================
 */

else {

    /*
     * ESTA es la única situación
     * donde aparece la pantalla azul.
     */

    actualizarDesdeMicrosoft(
        true
    );

}


/* =========================================
   ACTUALIZACIÓN AUTOMÁTICA
   ========================================= */

setInterval(() => {

    /*
     * SIEMPRE silencioso.
     *
     * Nunca muestra overlay.
     */

    actualizarDesdeMicrosoft(
        false
    );

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
                .querySelectorAll(
                    ".card"
                )
                .forEach(card => {

                    const contenido =
                        card.innerText
                            .toLowerCase();


                    card.style.display =
                        contenido.includes(
                            texto
                        )
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
        .querySelectorAll(
            ".card"
        )
        .forEach(card => {

            const estado =
                card.dataset.estado;


            let mostrar =
                false;


            switch (tipo) {

                case "Available":

                    mostrar =
                        estado ===
                        "Available";

                    break;


                case "Busy":

                    mostrar =
                        estado === "Busy" ||

                        estado ===
                        "InAMeeting" ||

                        estado ===
                        "OnACall";

                    break;


                case "Away":

                    mostrar =
                        estado === "Away" ||

                        estado ===
                        "BeRightBack";

                    break;


                case "Offline":

                    mostrar =
                        estado ===
                        "Offline";

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
        .querySelectorAll(
            ".card"
        )
        .forEach(card => {

            card.style.display =
                "";

        });

}