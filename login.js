/* =========================================
   CONFIGURACIÓN MSAL
========================================= */

const ALFERZA_TENANT_ID =
    "dbab984f-4bb1-4b60-9dff-da59f54acdf1";


const msalConfig = {

    auth: {

        clientId:
            "5d98417c-74a7-4fab-8f2c-41ac127be696",

        /*
         * SOLO TENANT ALFERZA
         */
        authority:
            `https://login.microsoftonline.com/${ALFERZA_TENANT_ID}`,

        redirectUri:
            "https://AlferzaTI.github.io/live-office/blank.html"

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
   VALIDAR CUENTA ALFERZA
========================================= */

function esCuentaAlferza(account) {

    if (!account) {

        return false;

    }


    /*
     * tid = Tenant ID de Microsoft Entra
     *
     * Solo aceptamos cuentas pertenecientes
     * al tenant oficial de ALFERZA.
     */

    const tenantId =
        account.tenantId ||
        account.idTokenClaims?.tid ||
        "";


    return (
        tenantId.toLowerCase() ===
        ALFERZA_TENANT_ID.toLowerCase()
    );

}


/* =========================================
   OBTENER PÁGINA DE DESTINO
========================================= */

function obtenerPaginaDestino() {

    const pagina =
        sessionStorage.getItem(
            "alferza_return_url"
        );


    if (
        pagina &&
        pagina.startsWith("/live-office/")
    ) {

        sessionStorage.removeItem(
            "alferza_return_url"
        );


        return pagina;

    }


    return "/live-office/index.html";

}

/* =========================================
   RECHAZAR CUENTA NO AUTORIZADA
========================================= */

async function rechazarCuenta(account) {

    console.warn(
        "Cuenta rechazada:",
        account?.username
    );


    sessionStorage.removeItem(
        "alferza_login"
    );


    mostrarMensaje(
        "Esta cuenta no pertenece a ALFERZA.",
        "error"
    );


    try {

        await msalInstance.logoutPopup({

            account:
                account

        });

    }

    catch (error) {

        console.warn(
            "No se pudo cerrar la sesión:",
            error
        );

    }


    habilitarBoton();

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
   ENTRAR AL SISTEMA
========================================= */

function entrarAlSistema() {

    sessionStorage.setItem(
        "alferza_login",
        "true"
    );


    const paginaDestino =
        obtenerPaginaDestino();


    window.location.replace(
        paginaDestino
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
         * VALIDACIÓN DEL TENANT
         */

        if (!esCuentaAlferza(cuenta)) {

            await rechazarCuenta(
                cuenta
            );

            return;

        }


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


        sessionStorage.removeItem(
            "alferza_login"
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
                    scopes,

                /*
                 * Muestra selección de cuenta.
                 */

                prompt:
                    "select_account"

            });


        const cuenta =
            respuesta.account;


        if (!cuenta) {

            throw new Error(
                "Microsoft no devolvió una cuenta."
            );

        }


        /* =====================================
           VALIDACIÓN ALFERZA
        ===================================== */

        if (!esCuentaAlferza(cuenta)) {

            await rechazarCuenta(
                cuenta
            );

            return;

        }


        msalInstance.setActiveAccount(
            cuenta
        );


        mostrarMensaje(
            "Validando acceso..."
        );


        /*
         * Obtener token y comprobar
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


        sessionStorage.setItem(
            "alferza_login",
            "true"
        );


        setTimeout(() => {

            entrarAlSistema();

        }, 400);

    }

    catch (error) {

        console.error(
            "Error de autenticación:",
            error
        );


        sessionStorage.removeItem(
            "alferza_login"
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