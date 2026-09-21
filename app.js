/* =========================================
   CONFIGURACIÓN MSAL (Azure AD - ALFERZA)
   ========================================= */

const msalConfig = {
    auth: {
        clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
        authority: "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",
        redirectUri: "https://alvaroalferza.github.io/live-office/blank.html"
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
const msalInstance = new msal.PublicClientApplication(msalConfig);

/* Pide el token: usa una sesión existente si la hay,
   o abre el popup de login la primera vez. */
async function obtenerToken() {

    let cuenta = msalInstance.getAllAccounts()[0];

    if (!cuenta) {
        const loginResponse = await msalInstance.loginPopup({ scopes });
        cuenta = loginResponse.account;
    }

    try {
        const response = await msalInstance.acquireTokenSilent({
            scopes,
            account: cuenta
        });
        return response.accessToken;

    } catch (error) {
        // Si el token silencioso falla (expiró la sesión, etc.), reintenta con popup
        const response = await msalInstance.acquireTokenPopup({ scopes });
        return response.accessToken;
    }
}


/* =========================================
   LÓGICA ORIGINAL (igual que antes, solo
   cambia de dónde sale el TOKEN)
   ========================================= */

async function cargarUsuarios() {

    const overlay = document.getElementById("loadingOverlay");
    overlay.classList.remove("oculto");

    const contenedor = document.getElementById("officeGrid");
    contenedor.innerHTML = "";

    const TOKEN = await obtenerToken();

    const respuesta = await fetch(
        "https://graph.microsoft.com/v1.0/users?$top=999",
        {
            headers: {
                Authorization: `Bearer ${TOKEN}`
            }
        }
    );

    const data = await respuesta.json();

    const usuarios = data.value.filter(
        u =>
            u.mail &&
            u.mail.toLowerCase().endsWith("@alferza.pe")
    );

    let disponibles = 0;
    let ocupados = 0;
    let ausentes = 0;
    let offline = 0;

    /* Una sola llamada trae la presencia de todos (hasta 650 ids por request) */

    const idsUsuarios = usuarios.map(u => u.id);

    let presenciaPorId = {};

    try {

        const presenciaResponse = await fetch(
            "https://graph.microsoft.com/v1.0/communications/getPresencesByUserId",
            {
                method: "POST",
                headers: {
                    Authorization: `Bearer ${TOKEN}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({ ids: idsUsuarios })
            }
        );

        const presenciaData = await presenciaResponse.json();

        (presenciaData.value || []).forEach(p => {
            presenciaPorId[p.id] = p;
        });

    } catch {
        // Si falla el batch, todos quedan como Offline por defecto (fallback abajo)
    }

    const usuariosConPresencia = usuarios.map(usuario => ({
        usuario,
        presencia: presenciaPorId[usuario.id] || { availability: "Offline" }
    }));

    let html = "";

    usuariosConPresencia.forEach(item => {

        const usuario = item.usuario;
        const presencia = item.presencia;

        let estado = presencia.availability || "Offline";

        let clase = "offline";

        switch (estado) {

            case "Available":
                clase = "disponible";
                disponibles++;
                break;

            case "Busy":
            case "InAMeeting":
            case "OnACall":
                clase = "ocupado";
                ocupados++;
                break;

            case "Away":
            case "BeRightBack":
                clase = "ausente";
                ausentes++;
                break;

            default:
                clase = "offline";
                offline++;
                break;
        }

        html += `
            <div class="card" data-estado="${estado}">
                <h3>${usuario.displayName}</h3>
                <p>${usuario.mail}</p>
                <div class="status ${clase}">
                    ${estado}
                </div>
            </div>
        `;

    });

    contenedor.innerHTML = html;

    document.getElementById("disp").innerText = disponibles;
    document.getElementById("busy").innerText = ocupados;
    document.getElementById("away").innerText = ausentes;
    document.getElementById("offline").innerText = offline;

    overlay.classList.add("oculto");
}

cargarUsuarios();

/* Actualizar cada 5 minutos */

setInterval(() => {
    cargarUsuarios();
}, 300000);

/* BUSCADOR */

document
    .getElementById("buscador")
    .addEventListener("keyup", function () {

        const texto = this.value.toLowerCase();

        document
            .querySelectorAll(".card")
            .forEach(card => {

                const contenido =
                    card.innerText.toLowerCase();

                card.style.display =
                    contenido.includes(texto)
                        ? ""
                        : "none";

            });

    });

/* FILTROS */

function filtrarEstado(tipo) {

    document
        .querySelectorAll(".card")
        .forEach(card => {

            const estado =
                card.dataset.estado;

            let mostrar = false;

            switch (tipo) {

                case "Available":
                    mostrar =
                        estado === "Available";
                    break;

                case "Busy":
                    mostrar =
                        estado === "Busy" ||
                        estado === "InAMeeting" ||
                        estado === "OnACall";
                    break;

                case "Away":
                    mostrar =
                        estado === "Away" ||
                        estado === "BeRightBack";
                    break;

                case "Offline":
                    mostrar =
                        estado === "Offline";
                    break;

                default:
                    mostrar = true;
            }

            card.style.display =
                mostrar ? "" : "none";

        });

}

function mostrarTodos() {

    document
        .querySelectorAll(".card")
        .forEach(card => {

            card.style.display = "";

        });

}