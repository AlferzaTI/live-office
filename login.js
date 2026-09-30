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

        /*
         * VERCEL
         */
        redirectUri:
            "https://liveofficealfeza.vercel.app/blank.html"

    },

    cache: {

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


/* =========================================
   PERMISOS
========================================= */

const scopes = [

    "User.Read",

    "Presence.Read.All",

    "Sites.Read.All",

    "AuditLog.Read.All",

    "DeviceManagementManagedDevices.Read.All"

];


/* =========================================
   MSAL
========================================= */

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

    if (!loginStatus) {
        return;
    }

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

    if (loginButton) {

        loginButton.disabled =
            true;

    }

    if (loginButtonText) {

        loginButtonText.textContent =
            mensaje;

    }

}


function habilitarBoton() {

    if (loginButton) {

        loginButton.disabled =
            false;

    }

    if (loginButtonText) {

        loginButtonText.textContent =
            "Iniciar sesión con Microsoft";

    }

}


/* =========================================
   VALIDAR CUENTA ALFERZA
========================================= */

function esCuentaAlferza(account) {

    if (!account) {

        return false;

    }


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


    /*
     * VERCEL
     *
     * Ya no usamos /live-office/
     */
    if (
        pagina &&
        pagina.startsWith("/")
    ) {

        sessionStorage.removeItem(
            "alferza_return_url"
        );


        /*
         * Evitar volver a login.html
         */
        if (
            !pagina.endsWith(
                "/login.html"
            )
        ) {

            return pagina;

        }

    }


    /*
     * Página principal de Vercel
     */
    return "/index.html";

}


/* =========================================
   RECHAZAR CUENTA NO AUTORIZADA
========================================= */

async function rechazarCuenta(
    account
) {

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
   OBTENER CUENTA
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


    console.log(
        "Redirigiendo a:",
        paginaDestino
    );


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
         * Intentamos recuperar el token
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

        console.error(
            "Error recuperando sesión:",
            error
        );


        console.error(
            "Código:",
            error?.errorCode
        );


        console.error(
            "Mensaje:",
            error?.errorMessage ||
            error?.message
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
                 * Permite seleccionar cuenta
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


        console.log(
            "Cuenta:",
            cuenta.username
        );


        console.log(
            "Tenant:",
            cuenta.tenantId
        );


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
         * Obtener token
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
            "================================="
        );

        console.error(
            "ERROR DE AUTENTICACIÓN"
        );

        console.error(
            "================================="
        );

        console.error(
            "Código:",
            error?.errorCode
        );

        console.error(
            "Mensaje:",
            error?.errorMessage ||
            error?.message
        );

        console.error(
            "Nombre:",
            error?.name
        );

        console.error(
            "Error completo:",
            error
        );


        sessionStorage.removeItem(
            "alferza_login"
        );


        habilitarBoton();


        /*
         * Mostrar el error real temporalmente
         * para poder detectar el problema.
         */

        const mensajeError =
            error?.errorMessage ||
            error?.message ||
            "Error desconocido";


        mostrarMensaje(
            mensajeError,
            "error"
        );

    }

}


/* =========================================
   EVENTO BOTÓN
========================================= */

if (loginButton) {

    loginButton.addEventListener(
        "click",
        iniciarSesion
    );

}


/* =========================================
   INICIO
========================================= */

comprobarSesion();