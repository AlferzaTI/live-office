const msalConfig = {
    auth: {
        clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
        authority: "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",
        redirectUri: "https://alvaroalferza.github.io/live-office/blank.html"
    }
};

const scopes = [
    "User.Read",
    "Presence.Read.All",
    "Sites.Read.All"
];

const msalInstance =
    new msal.PublicClientApplication(msalConfig);

let inicializado = false;


async function inicializarMSAL() {

    if (!inicializado) {

        await msalInstance.initialize();

        inicializado = true;

    }

}


export async function obtenerToken() {

    await inicializarMSAL();


    let cuenta =
        msalInstance.getAllAccounts()[0];


    if (!cuenta) {

        const loginResponse =
            await msalInstance.loginPopup({
                scopes
            });

        cuenta =
            loginResponse.account;

    }


    try {

        const respuesta =
            await msalInstance.acquireTokenSilent({
                scopes,
                account: cuenta
            });

        return respuesta.accessToken;

    }
    catch (error) {

        console.warn(
            "No se pudo obtener el token silenciosamente.",
            error
        );


        const respuesta =
            await msalInstance.acquireTokenPopup({
                scopes
            });


        return respuesta.accessToken;

    }

}