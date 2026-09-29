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


const msalInstance =
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
        msalInstance.getAllAccounts();


    if (!cuentas.length) {

        throw new Error(
            "No se encontró una cuenta Microsoft activa."
        );

    }


    const cuenta =
        cuentas[0];


    msalInstance.setActiveAccount(
        cuenta
    );


    const respuesta =
        await msalInstance.acquireTokenSilent({

            scopes: [
                "User.Read",
                "Sites.Read.All"
            ],

            account:
                cuenta

        });


    return respuesta.accessToken;

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
            "Error SharePoint Site:",
            error
        );

        throw new Error(
            `SharePoint Site: HTTP ${respuesta.status}`
        );

    }


    return await respuesta.json();

}


/* =========================================
   OBTENER LISTA MONITOREOTI
========================================= */

async function obtenerMonitoreoTI() {

    try {

        console.log(
            "🔄 Conectando con Microsoft..."
        );


        const TOKEN =
            await obtenerTokenInfraestructura();


        console.log(
            "✅ Token obtenido"
        );


        const sitio =
            await obtenerSitioSharePoint(
                TOKEN
            );


        console.log(
            "✅ Sitio SharePoint encontrado:",
            sitio
        );


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
                "Error obteniendo lista:",
                error
            );

            throw new Error(
                `MonitoreoTI: HTTP ${respuesta.status}`
            );

        }


        const data =
            await respuesta.json();


        console.log(
            "================================="
        );

        console.log(
            "✅ MONITOREO TI"
        );

        console.log(
            "================================="
        );

        console.table(
            data.value
        );


        return data.value;

    }

    catch (error) {

        console.error(
            "❌ Error Monitoreo TI:",
            error
        );

        return [];

    }

}


/* =========================================
   INICIAR
========================================= */

async function iniciarMonitoreo() {

    console.log(
        "🚀 Iniciando Centro de Monitoreo TI..."
    );


    const datos =
        await obtenerMonitoreoTI();


    console.log(
        "Datos recibidos:",
        datos
    );

}


iniciarMonitoreo();