/* =========================================
   INFRAESTRUCTURA - MONITOREO TI
========================================= */


/* =========================================
   CONFIGURACIÓN MICROSOFT
========================================= */

const msalConfigInfraestructura = {

    auth: {

        clientId:
            "5d98417c-74a7-4fab-8f2c-41ac127be696",

        authority:
            "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",

        redirectUri:
            "https://alferzati.github.io/live-office/blank.html"

    },

    cache: {

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


/* =========================================
   INSTANCIA MSAL
========================================= */

const msalInstanceInfraestructura =
    new msal.PublicClientApplication(
        msalConfigInfraestructura
    );


/* =========================================
   CONFIGURACIÓN SHAREPOINT
========================================= */

const SHAREPOINT_HOSTNAME =
    "alferzaholding-my.sharepoint.com";

const SHAREPOINT_SITE =
    "/personal/soporte1_alferza_pe";

const SHAREPOINT_LIST =
    "MonitoreoTI";


/* =========================================
   OBTENER TOKEN
========================================= */

async function obtenerTokenInfraestructura() {

    const cuentas =
        msalInstanceInfraestructura
            .getAllAccounts();


    if (!cuentas.length) {

        throw new Error(
            "No se encontró una cuenta Microsoft activa."
        );

    }


    const cuenta =
        cuentas[0];


    msalInstanceInfraestructura
        .setActiveAccount(cuenta);


    try {

        const respuesta =
            await msalInstanceInfraestructura
                .acquireTokenSilent({

                    scopes: [

                        "User.Read",

                        "Sites.Read.All"

                    ],

                    account: cuenta

                });


        console.log(
            "✅ Token obtenido correctamente"
        );


        return respuesta.accessToken;

    }

    catch (error) {

        console.error(
            "❌ Error obteniendo token:",
            error
        );

        throw error;

    }

}


/* =========================================
   OBTENER SITIO SHAREPOINT
========================================= */

async function obtenerSitioSharePoint(
    TOKEN
) {

    const respuesta =
        await fetch(

            `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOSTNAME}:${SHAREPOINT_SITE}`,

            {

                method: "GET",

                headers: {

                    Authorization:
                        `Bearer ${TOKEN}`

                }

            }

        );


    if (!respuesta.ok) {

        const error =
            await respuesta.text();


        console.error(
            "❌ Error obteniendo sitio SharePoint:",
            error
        );


        throw new Error(
            `SharePoint Site: HTTP ${respuesta.status}`
        );

    }


    const sitio =
        await respuesta.json();


    console.log(
        "✅ Sitio SharePoint encontrado:",
        sitio
    );


    return sitio;

}


/* =========================================
   OBTENER LISTA MONITOREOTI
========================================= */

async function obtenerMonitoreoTI() {

    try {

        console.log(
            "🔄 Conectando con Microsoft Graph..."
        );


        /*
         * TOKEN
         */

        const TOKEN =
            await obtenerTokenInfraestructura();


        /*
         * SITIO
         */

        const sitio =
            await obtenerSitioSharePoint(
                TOKEN
            );


        /*
         * LISTA
         */

        const respuesta =
            await fetch(

                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${SHAREPOINT_LIST}/items?expand=fields`,

                {

                    method: "GET",

                    headers: {

                        Authorization:
                            `Bearer ${TOKEN}`

                    }

                }

            );


        if (!respuesta.ok) {

            const error =
                await respuesta.text();


            console.error(
                "❌ Error obteniendo lista MonitoreoTI:",
                error
            );


            throw new Error(
                `MonitoreoTI: HTTP ${respuesta.status}`
            );

        }


        const data =
            await respuesta.json();


        console.log(
            "===================================="
        );

        console.log(
            "✅ DATOS DE MONITOREO TI"
        );

        console.log(
            "===================================="
        );


        console.table(
            data.value
        );


        return data.value;

    }

    catch (error) {

        console.error(
            "❌ Error obteniendo Monitoreo TI:",
            error
        );


        return [];

    }

}


/* =========================================
   INICIAR MONITOREO
========================================= */

async function iniciarMonitoreo() {

    console.log(
        "🚀 Iniciando Centro de Monitoreo TI..."
    );


    const datos =
        await obtenerMonitoreoTI();


    console.log(
        "Monitoreo TI cargado:",
        datos
    );

}


iniciarMonitoreo();