import { obtenerToken } from "./salas-api.js";

/* =========================================================
   ALFERZA LIVE OFFICE
   MICROSOFT LISTS - RESERVA DE SALAS
========================================================= */


/* =========================================================
   OBTENER INFORMACIÓN DE MICROSOFT LISTS
========================================================= */

async function probarMicrosoftLists() {

    try {

        console.log("======================================");
        console.log("ALFERZA LIVE OFFICE");
        console.log("MICROSOFT LISTS");
        console.log("======================================");


        /*
        ======================================
        0. OBTENER TOKEN
        ======================================
        */

        const token = await obtenerToken();

        console.log("TOKEN OBTENIDO");


        /*
        ======================================
        1. OBTENER SITE DE SHAREPOINT
        ======================================
        */

        const siteUrl =
            "https://graph.microsoft.com/v1.0/sites/alferzaholding-my.sharepoint.com:/personal/soporte1_alferza_pe";

        console.log("======================================");
        console.log("CONSULTANDO SHAREPOINT SITE");
        console.log("======================================");

        console.log(siteUrl);


        const respuestaSite =
            await fetch(
                siteUrl,
                {
                    method: "GET",

                    headers: {
                        Authorization: `Bearer ${token}`,
                        Accept: "application/json"
                    }
                }
            );


        if (!respuestaSite.ok) {

            const error =
                await respuestaSite.text();

            throw new Error(
                `Error obteniendo SharePoint Site (${respuestaSite.status}): ${error}`
            );

        }


        const site =
            await respuestaSite.json();


        console.log("======================================");
        console.log("SITE ENCONTRADO");
        console.log("======================================");

        console.log(site);


        console.log("SITE ID:");
        console.log(site.id);


        /*
        ======================================
        2. OBTENER LISTA ReservaSalas
        ======================================
        */

        const listaUrl =
            `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/ReservaSalas?expand=columns`;


        console.log("======================================");
        console.log("CONSULTANDO LISTA ReservaSalas");
        console.log("======================================");

        console.log(listaUrl);


        const respuestaLista =
            await fetch(
                listaUrl,
                {
                    method: "GET",

                    headers: {
                        Authorization: `Bearer ${token}`,
                        Accept: "application/json"
                    }
                }
            );


        if (!respuestaLista.ok) {

            const error =
                await respuestaLista.text();

            throw new Error(
                `Error obteniendo ReservaSalas (${respuestaLista.status}): ${error}`
            );

        }


        const lista =
            await respuestaLista.json();


        console.log("======================================");
        console.log("LISTA ENCONTRADA");
        console.log("======================================");

        console.log(lista);


        /*
        ======================================
        3. INFORMACIÓN DE LA LISTA
        ======================================
        */

        console.log("======================================");
        console.log("INFORMACIÓN DE LA LISTA");
        console.log("======================================");

        console.log("ID:");
        console.log(lista.id);

        console.log("NOMBRE:");
        console.log(lista.name);

        console.log("DISPLAY NAME:");
        console.log(lista.displayName);


        /*
        ======================================
        4. COLUMNAS
        ======================================
        */

        console.log("======================================");
        console.log("COLUMNAS DE ReservaSalas");
        console.log("======================================");


        if (
            lista.columns &&
            Array.isArray(lista.columns)
        ) {

            lista.columns.forEach(
                (columna, index) => {

                    console.log(
                        `${index + 1}.`
                    );

                    console.log(
                        "name:",
                        columna.name
                    );

                    console.log(
                        "displayName:",
                        columna.displayName
                    );

                    console.log(
                        "description:",
                        columna.description
                    );

                    console.log(
                        "type:",
                        columna.columnGroup
                    );

                    console.log(
                        "------------------------------"
                    );

                }
            );

        }
        else {

            console.log(
                "No se encontraron columnas."
            );

        }


        /*
        ======================================
        5. OBTENER REGISTROS
        ======================================
        */

        const itemsUrl =
            `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/${lista.id}/items?expand=fields`;


        console.log("======================================");
        console.log("CONSULTANDO REGISTROS");
        console.log("======================================");

        console.log(itemsUrl);


        const respuestaItems =
            await fetch(
                itemsUrl,
                {
                    method: "GET",

                    headers: {
                        Authorization: `Bearer ${token}`,
                        Accept: "application/json"
                    }
                }
            );


        if (!respuestaItems.ok) {

            const error =
                await respuestaItems.text();

            throw new Error(
                `Error obteniendo registros (${respuestaItems.status}): ${error}`
            );

        }


        const datosItems =
            await respuestaItems.json();


        /*
        ======================================
        6. INFORMACIÓN DE REGISTROS
        ======================================
        */

        console.log("======================================");
        console.log("REGISTROS DE ReservaSalas");
        console.log("======================================");


        const items =
            datosItems.value || [];


        console.log(
            `Cantidad de registros: ${items.length}`
        );


        /*
        ======================================
        7. MOSTRAR CADA REGISTRO
        ======================================
        */

        items.forEach(
            (item, index) => {

                console.log(
                    "======================================"
                );

                console.log(
                    `REGISTRO ${index + 1}`
                );

                console.log(
                    "ID:",
                    item.id
                );

                console.log(
                    "ITEM COMPLETO:",
                    item
                );

                console.log(
                    "FIELDS:"
                );

                console.log(
                    item.fields
                );

            }
        );


        /*
        ======================================
        8. MOSTRAR TODOS LOS FIELDS
        ======================================
        */

        console.log("======================================");
        console.log("TODOS LOS FIELDS");
        console.log("======================================");


        items.forEach(
            (item, index) => {

                console.log(
                    `FIELDS REGISTRO ${index + 1}:`,
                    item.fields
                );

            }
        );


        /*
        ======================================
        9. PAGINACIÓN
        ======================================
        */

        if (datosItems["@odata.nextLink"]) {

            console.log("======================================");
            console.log("HAY MÁS REGISTROS");
            console.log("======================================");

            console.log(
                datosItems["@odata.nextLink"]
            );

        }
        else {

            console.log("======================================");
            console.log("NO HAY MÁS REGISTROS");
            console.log("======================================");

        }


        /*
        ======================================
        10. RESUMEN FINAL
        ======================================
        */

        console.log("======================================");
        console.log("RESUMEN");
        console.log("======================================");

        console.log(
            "Site ID:",
            site.id
        );

        console.log(
            "Lista:",
            lista.displayName || lista.name
        );

        console.log(
            "Lista ID:",
            lista.id
        );

        console.log(
            "Cantidad de registros:",
            items.length
        );

        console.log("======================================");
        console.log("FIN DE CONSULTA");
        console.log("======================================");

    }
    catch (error) {

        console.error("======================================");
        console.error("ERROR");
        console.error("======================================");

        console.error(error);

    }

}


/* =========================================================
   INICIAR
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        probarMicrosoftLists();

    }
);