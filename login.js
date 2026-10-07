/* =========================================================
   ALFERZA LIVE OFFICE
   LOGIN
========================================================= */


/* =========================================================
   CONFIGURACIÓN
========================================================= */

const ALFERZA_TENANT_ID =
    "dbab984f-4bb1-4b60-9dff-da59f54acdf1";

const ALFERZA_CLIENT_ID =
    "5d98417c-74a7-4fab-8f2c-41ac127be696";

const ALFERZA_REDIRECT_URI =
    "https://alferzati.github.io/live-office/blank.html";


const msalConfig = {

    auth: {

        clientId:
            ALFERZA_CLIENT_ID,

        authority:
            `https://login.microsoftonline.com/${ALFERZA_TENANT_ID}`,

        redirectUri:
            ALFERZA_REDIRECT_URI

    },

    cache: {

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


/* =========================================================
   SCOPES
========================================================= */

/*
 * Para iniciar sesión NO necesitamos pedir todos los
 * permisos administrativos de Seguridad.
 *
 * Primero autenticamos al usuario.
 *
 * Los permisos adicionales se solicitan después mediante
 * acquireTokenSilent() desde auth.js.
 */

const loginScopes = [
    "User.Read"
];


/* =========================================================
   MSAL
========================================================= */

const msalInstance =
    new msal.PublicClientApplication(
        msalConfig
    );


/* =========================================================
   ELEMENTOS
========================================================= */

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


/* =========================================================
   MENSAJES
========================================================= */

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


/* =========================================================
   BOTÓN
========================================================= */

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


/* =========================================================
   VALIDAR TENANT ALFERZA
========================================================= */

function esCuentaAlferza(
    account
) {

    if (!account) {

        return false;

    }


    const tenantId =
        account.tenantId ||
        account.idTokenClaims?.tid ||
        "";


    return (
        String(tenantId).toLowerCase() ===
        ALFERZA_TENANT_ID.toLowerCase()
    );

}


/* =========================================================
   OBTENER CUENTA ACTUAL
========================================================= */

function obtenerCuenta() {

    const cuentaActiva =
        msalInstance.getActiveAccount();


    if (cuentaActiva) {

        return cuentaActiva;

    }


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


/* =========================================================
   OBTENER DESTINO
========================================================= */

function obtenerPaginaDestino() {

    const pagina =
        sessionStorage.getItem(
            "alferza_return_url"
        );


    if (
        pagina &&
        pagina.startsWith(
            "/live-office/"
        ) &&
        !pagina.includes(
            "/login.html"
        ) &&
        !pagina.includes(
            "/blank.html"
        )
    ) {

        sessionStorage.removeItem(
            "alferza_return_url"
        );

        return pagina;

    }


    return "/live-office/index.html";

}


/* =========================================================
   LIMPIAR ESTADO LOCAL
========================================================= */

function limpiarEstadoLocal() {

    sessionStorage.removeItem(
        "alferza_login"
    );

    sessionStorage.removeItem(
        "alferza_permisos"
    );

}


/* =========================================================
   RECHAZAR CUENTA
========================================================= */

async function rechazarCuenta(
    account
) {

    console.warn(
        "Cuenta rechazada:",
        account?.username
    );


    limpiarEstadoLocal();


    mostrarMensaje(
        "Esta cuenta no pertenece a ALFERZA.",
        "error"
    );


    try {

        await msalInstance.logoutPopup({

            account:
                account,

            postLogoutRedirectUri:
                window.location.origin +
                "/live-office/login.html"

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


/* =========================================================
   OBTENER TOKEN
========================================================= */

async function obtenerToken(
    account
) {

    return await msalInstance.acquireTokenSilent({

        scopes:
            loginScopes,

        account:
            account

    });

}


/* =========================================================
   ENTRAR AL SISTEMA
========================================================= */

function entrarAlSistema() {

    sessionStorage.setItem(
        "alferza_login",
        "true"
    );


    const paginaDestino =
        obtenerPaginaDestino();


    console.log(
        "Destino:",
        paginaDestino
    );


    window.location.replace(
        paginaDestino
    );

}


/* =========================================================
   COMPROBAR SESIÓN EXISTENTE
========================================================= */

async function comprobarSesion() {

    try {

        const cuenta =
            obtenerCuenta();


        /*
         * No existe sesión MSAL.
         */

        if (!cuenta) {

            limpiarEstadoLocal();

            habilitarBoton();

            return;

        }


        console.log(
            "Cuenta MSAL encontrada:",
            cuenta.username
        );


        /*
         * Validar tenant.
         */

        if (
            !esCuentaAlferza(
                cuenta
            )
        ) {

            await rechazarCuenta(
                cuenta
            );

            return;

        }


        /*
         * Establecer cuenta activa.
         */

        msalInstance.setActiveAccount(
            cuenta
        );


        mostrarMensaje(
            "Sesión encontrada. Validando..."
        );


        /*
         * Comprobar que MSAL puede obtener
         * el token silenciosamente.
         */

        await obtenerToken(
            cuenta
        );


        /*
         * Todo correcto.
         */

        mostrarMensaje(
            "Sesión válida. Ingresando...",
            "success"
        );


        sessionStorage.setItem(
            "alferza_login",
            "true"
        );


        setTimeout(
            () => {

                entrarAlSistema();

            },
            300
        );

    }

    catch (error) {

        console.warn(
            "No se pudo recuperar la sesión:",
            error
        );


        /*
         * Si la sesión MSAL ya no es válida,
         * limpiamos nuestro indicador.
         */

        limpiarEstadoLocal();


        habilitarBoton();


        mostrarMensaje(
            "Inicia sesión para continuar."
        );

    }

}


/* =========================================================
   INICIAR SESIÓN
========================================================= */

async function iniciarSesion() {

    try {

        bloquearBoton(
            "Conectando con Microsoft..."
        );


        mostrarMensaje(
            "Abriendo Microsoft..."
        );


        /*
         * IMPORTANTE:
         *
         * Ya NO usamos select_account.
         *
         * Si Microsoft ya tiene una cuenta válida,
         * intentará utilizarla directamente.
         */

        const respuesta =
            await msalInstance.loginPopup({

                scopes:
                    loginScopes,

                prompt:
                    "login"

            });


        const cuenta =
            respuesta?.account;


        if (!cuenta) {

            throw new Error(
                "Microsoft no devolvió una cuenta."
            );

        }


        console.log(
            "Usuario autenticado:",
            cuenta.username
        );


        /* =========================================
           VALIDAR TENANT
        ========================================= */

        if (
            !esCuentaAlferza(
                cuenta
            )
        ) {

            await rechazarCuenta(
                cuenta
            );

            return;

        }


        /* =========================================
           CUENTA ACTIVA
        ========================================= */

        msalInstance.setActiveAccount(
            cuenta
        );


        mostrarMensaje(
            "Validando acceso..."
        );


        /* =========================================
           VALIDAR TOKEN
        ========================================= */

        await obtenerToken(
            cuenta
        );


        /* =========================================
           SESIÓN LOCAL
        ========================================= */

        sessionStorage.setItem(
            "alferza_login",
            "true"
        );


        mostrarMensaje(
            "Acceso autorizado. Ingresando...",
            "success"
        );


        setTimeout(
            () => {

                entrarAlSistema();

            },
            400
        );

    }

    catch (error) {

        console.error(
            "ERROR DE AUTENTICACIÓN:",
            error
        );


        limpiarEstadoLocal();


        habilitarBoton();


        /*
         * Mensajes más útiles para detectar
         * exactamente qué está fallando.
         */

        const codigo =
            error?.errorCode ||
            "";


        if (
            codigo ===
            "user_cancelled"
        ) {

            mostrarMensaje(
                "Inicio de sesión cancelado.",
                "error"
            );

            return;

        }


        if (
            codigo ===
            "interaction_required"
        ) {

            mostrarMensaje(
                "Microsoft requiere volver a autenticar la cuenta.",
                "error"
            );

            return;

        }


        mostrarMensaje(
            "No se pudo iniciar sesión. Revisa la consola para ver el error.",
            "error"
        );

    }

}


/* =========================================================
   EVENTO BOTÓN
========================================================= */

if (loginButton) {

    loginButton.addEventListener(
        "click",
        iniciarSesion
    );

}


/* =========================================================
   INICIO
========================================================= */

comprobarSesion();