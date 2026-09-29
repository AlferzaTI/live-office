/* =========================================
   INFRAESTRUCTURA - MONITOREO TI
========================================= */


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
   OBTENER TOKEN MICROSOFT
========================================= */

async function obtenerTokenInfraestructura() {

    /*
     * Utilizamos la misma instancia MSAL
     * que ya utiliza ALFERZA LIVE OFFICE.
     */

    const cuentas =
        msalInstance.getAllAccounts();


    const cuenta =
        cuentas[0];


    if (!cuenta) {

        throw new Error(
            "No existe una sesión de Microsoft."
        );

    }


    msalInstance.setActiveAccount(
        cuenta
    );


    try {

        const response =
            await msalInstance.acquireTokenSilent({

                scopes: [

                    "User.Read",

                    "Sites.Read.All"

                ],

                account:
                    cuenta

            });


        return response.accessToken;

    }

    catch (error) {

        console.error(
            "Error obteniendo token para Infraestructura:",
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

        throw new Error(

            `SharePoint Site: HTTP ${respuesta.status}`

        );

    }


    return await respuesta.json();

}


/* =========================================
   OBTENER MONITOREO TI
========================================= */

async function obtenerMonitoreoTI() {

    try {

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


        console.log(
            "Sitio SharePoint:",
            sitio
        );


        /*
         * LISTA MONITOREOTI
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

            throw new Error(

                `SharePoint MonitoreoTI: HTTP ${respuesta.status}`

            );

        }


        const data =
            await respuesta.json();


        console.log(
            "DATOS MONITOREO TI:",
            data.value
        );


        return data.value;

    }

    catch (error) {

        console.error(
            "Error obteniendo Monitoreo TI:",
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
        "Iniciando Centro de Monitoreo TI..."
    );


    const datos =
        await obtenerMonitoreoTI();


    console.log(
        "Monitoreo TI cargado:",
        datos
    );

}


/* =========================================
   INICIO
========================================= */

iniciarMonitoreo();