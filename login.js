/* =========================================
   CONFIGURACIÓN MSAL
========================================= */

const msalConfig = {

    auth: {

        clientId:
            "5d98417c-74a7-4fab-8f2c-41ac127be696",

        authority:
            "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",

        redirectUri:
            "https://alvaroalferza.github.io/live-office/blank.html"

    },

    cache: {

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


const scopes = [

    "User.Read",

    "Presence.Read.All",

    "Sites.Read.All"

];


const msalInstance =
    new msal.PublicClientApplication(
        msalConfig
    );


/* =========================================
   ELEMENTOS
========================================= */

const loginButton =
    document.getElementById(
        "loginButton"
    );


const loginButtonText =
    document.getElementById(
        "loginButtonText"
    );


const loginStatus =
    document.getElementById(
        "loginStatus"
    );


/* =========================================
   MENSAJE
========================================= */

function mostrarMensaje(
    mensaje,
    tipo = ""
) {

    loginStatus.textContent =
        mensaje;

    loginStatus.className =
        "login-status";


    if (tipo) {

        loginStatus.classList.add(
            tipo
        );

    }

}


/* =========================================
   BOTÓN
========================================= */

function bloquearBoton(
    mensaje
) {

    loginButton.disabled =
        true;

    loginButtonText.textContent =
        mensaje;

}


function habilitarBoton() {

    loginButton.disabled =
        false;

    loginButtonText.textContent =
        "Iniciar sesión con Microsoft";

}


/* =========================================
   CUENTA EXISTENTE
========================================= */

function obtenerCuenta() {

    const cuentas =
        msalInstance.getAllAccounts();


    if (
        cuentas &&
        cuentas.length > 0
    ) {

        return cuentas[0];

    }


    return null;

}


/* =========================================
   ENTRAR AL INDEX
========================================= */

function entrarAlSistema() {

    /*
     * Indicamos que el usuario ya pasó
     * correctamente por el login.
     */

    sessionStorage.setItem(
        "alferza_login",
        "true"
    );


    window.location.replace(
        "index.html"
    );

}


/* =========================================
   LOGIN AUTOMÁTICO
========================================= */

async function comprobarSesion() {

    try {

        const cuenta =
            obtenerCuenta();


        if (!cuenta) {

            habilitarBoton();

            return;

        }


        /*
         * Ya existe una cuenta MSAL.
         */

        msalInstance.setActiveAccount(
            cuenta
        );


        mostrarMensaje(
            "Sesión encontrada. Ingresando...",
            "success"
        );


        /*
         * Comprobamos silenciosamente
         * que el token siga disponible.
         */

        await msalInstance.acquireTokenSilent({

            scopes:
                scopes,

            account:
                cuenta

        });


        entrarAlSistema();

    }

    catch (error) {

        console.warn(
            "No se pudo recuperar la sesión:",
            error
        );


        habilitarBoton();

    }

}


/* =========================================
   INICIAR SESIÓN
========================================= */

async function iniciarSesion() {

    try {

        bloquearBoton(
            "Conectando con Microsoft..."
        );


        mostrarMensaje(
            "Abriendo Microsoft..."
        );


        const respuesta =
            await msalInstance.loginPopup({

                scopes:
                    scopes

            });


        const cuenta =
            respuesta.account;


        if (!cuenta) {

            throw new Error(
                "Microsoft no devolvió una cuenta."
            );

        }


        msalInstance.setActiveAccount(
            cuenta
        );


        mostrarMensaje(
            "Validando acceso..."
        );


        /*
         * Obtener token para comprobar
         * los permisos necesarios.
         */

        await msalInstance.acquireTokenSilent({

            scopes:
                scopes,

            account:
                cuenta

        });


        mostrarMensaje(
            "Acceso autorizado. Ingresando...",
            "success"
        );


        /*
         * Guardamos la sesión local
         * para que index.html sepa que
         * el usuario ya pasó por el login.
         */

        sessionStorage.setItem(
            "alferza_login",
            "true"
        );


        setTimeout(() => {

            window.location.replace(
                "index.html"
            );

        }, 400);

    }

    catch (error) {

        console.error(
            "Error de autenticación:",
            error
        );


        habilitarBoton();


        mostrarMensaje(
            "No se pudo iniciar sesión. Inténtalo nuevamente.",
            "error"
        );

    }

}


/* =========================================
   EVENTO BOTÓN
========================================= */

loginButton.addEventListener(
    "click",
    iniciarSesion
);


/* =========================================
   INICIO
========================================= */

comprobarSesion();