
(function () {
    "use strict";

    const CONFIG = {
        clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
        tenantId: "dbab984f-4bb1-4b60-9dff-da59f54acdf1",
        redirectUri: "https://alferzati.github.io/live-office/blank.html",
        basePath: "/live-office/",
        sharePointHost: "alferzaholding-my.sharepoint.com",
        sharePointPath: "/personal/soporte1_alferza_pe",
        listName: "PermisosTI",
        adminEmails: ["soporte1@alferza.pe"]
    };

    const MODULOS = [
        "Oficina",
        "Personal",
        "Reservas",
        "Salas",
        "Comunicados",
        "Seguridad",
        "Infraestructura",
        "Tickets",
        "Configuracion",
        "AsistenteTI"
    ];

    const PAGINAS = {
        "index.html": "Oficina",
        "personal.html": "Personal",
        "reserva.html": "Reservas",
        "reservas.html": "Reservas",
        "salas.html": "Salas",
        "comunicados.html": "Comunicados",
        "seguridad.html": "Seguridad",
        "infraestructura.html": "Infraestructura",
        "tickets.html": "Tickets",
        "chatbot.html": "AsistenteTI",
        "configuracion.html": "Configuracion"
    };

    const PAGINAS_PUBLICAS = ["login.html", "blank.html"];

    let instanciaMSAL = null;
    let tokenGraph = null;
    let cuentaActual = null;

    function normalizar(valor) {
        return String(valor ?? "")
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[\s_-]/g, "")
            .toLowerCase()
            .trim();
    }

    function normalizarCorreo(valor) {
        return String(valor ?? "").trim().toLowerCase();
    }

    function esVerdadero(valor) {
        if (valor === true || valor === 1) return true;

        return [
            "true", "1", "yes", "si", "sí", "verdadero"
        ].includes(String(valor ?? "").trim().toLowerCase());
    }

    function obtenerCampo(registro, nombres) {
        const claves = Object.keys(registro || {});

        for (const nombre of nombres) {
            const clave = claves.find(
                campo => normalizar(campo) === normalizar(nombre)
            );

            if (clave !== undefined) return registro[clave];
        }

        return undefined;
    }

    function mostrarDenegacion(mensaje, volverAlLogin, detalle) {
        function renderizar() {
            if (!document.body) return;

            document.body.innerHTML = `
                <main style="
                    min-height:100vh;
                    display:flex;
                    align-items:center;
                    justify-content:center;
                    padding:24px;
                    box-sizing:border-box;
                    background:#f4f6fa;
                    font-family:Arial,sans-serif;
                    color:#172033;
                    text-align:center;
                ">
                    <section style="
                        width:100%;
                        max-width:500px;
                        padding:36px 28px;
                        background:#fff;
                        border:1px solid #e1e6ef;
                        border-radius:16px;
                        box-sizing:border-box;
                    ">
                        <h1 style="color:#00205c">🔒 Acceso restringido</h1>

                        <p style="line-height:1.7;color:#596579">
                            ${mensaje}
                        </p>

                        ${detalle ? `
                            <p style="
                                padding:12px;
                                background:#f4f6fa;
                                border-radius:8px;
                                color:#596579;
                                font-size:13px;
                                overflow-wrap:anywhere;
                                text-align:left;
                            ">
                                <strong>Detalle:</strong>
                                ${detalle}
                            </p>
                        ` : ""}

                        <button id="alferza-access-back" type="button" style="
                            margin-top:16px;
                            padding:12px 20px;
                            border:0;
                            border-radius:8px;
                            background:#00205c;
                            color:white;
                            cursor:pointer;
                        ">
                            ${volverAlLogin
                                ? "Volver al inicio de sesión"
                                : "Volver a Oficina"}
                        </button>
                    </section>
                </main>
            `;

            document
                .getElementById("alferza-access-back")
                .addEventListener("click", function () {
                    if (volverAlLogin) {
                        sessionStorage.removeItem("alferza_login");
                        sessionStorage.removeItem("alferza_permisos");
                        sessionStorage.removeItem("alferza_account_username");

                        window.location.replace(
                            CONFIG.basePath + "login.html"
                        );
                    } else {
                        window.location.replace(
                            CONFIG.basePath + "index.html"
                        );
                    }
                });
        }

        if (document.body) {
            renderizar();
        } else {
            document.addEventListener("DOMContentLoaded", renderizar, {
                once: true
            });
        }
    }

    async function graph(endpoint) {
        if (!tokenGraph) {
            throw new Error("No se obtuvo un token de Microsoft Graph.");
        }

        const url = endpoint.startsWith("https://")
            ? endpoint
            : "https://graph.microsoft.com/v1.0" + endpoint;

        const respuesta = await fetch(url, {
            headers: {
                Authorization: "Bearer " + tokenGraph,
                Accept: "application/json"
            },
            cache: "no-store"
        });

        if (!respuesta.ok) {
            let detalle = "";

            try {
                const error = await respuesta.json();
                detalle = error.error?.message || "";
            } catch (_) {}

            throw new Error(
                "Microsoft Graph HTTP " + respuesta.status +
                (detalle ? ": " + detalle : "")
            );
        }

        return respuesta.json();
    }

    async function obtenerTodos(endpoint) {
        const resultados = [];
        let url = endpoint;

        while (url) {
            const pagina = await graph(url);
            resultados.push(...(pagina.value || []));
            url = pagina["@odata.nextLink"] || null;
        }

        return resultados;
    }

    async function obtenerListaPermisos() {
        const sitio = await graph(
            "/sites/" +
            CONFIG.sharePointHost +
            ":" +
            CONFIG.sharePointPath
        );

        const listas = await obtenerTodos(
            "/sites/" + sitio.id +
            "/lists?$select=id,name,displayName&$top=200"
        );

        const lista = listas.find(function (item) {
            return (
                normalizar(item.displayName) === normalizar(CONFIG.listName) ||
                normalizar(item.name) === normalizar(CONFIG.listName)
            );
        });

        if (!lista) {
            throw new Error(
                "No se encontró la lista '" + CONFIG.listName +
                "' en el sitio de SharePoint configurado."
            );
        }

        const registros = await obtenerTodos(
            "/sites/" + sitio.id +
            "/lists/" + lista.id +
            "/items?$expand=fields&$top=500"
        );

        console.info(
            "[ALFERZA ACCESO] Registros de permisos leídos:",
            registros.length
        );

        return registros.map(item => item.fields || {});
    }

    function buscarRegistro(registros, correos) {
        const correosBuscados = new Set(
            correos.map(normalizarCorreo).filter(Boolean)
        );

        return registros.find(function (registro) {
            const correoRegistrado = normalizarCorreo(
                obtenerCampo(registro, [
                    "UsuarioCorreo",
                    "Usuario Correo",
                    "Correo",
                    "Email",
                    "Mail"
                ])
            );

            return correosBuscados.has(correoRegistrado);
        });
    }

    async function obtenerPermisos(correos) {
        const registros = await obtenerListaPermisos();

        const registro = buscarRegistro(registros, correos);

        if (!registro) {
            console.warn(
                "[ALFERZA ACCESO] No se encontró coincidencia en PermisosTI.",
                { correosBuscados: correos }
            );

            throw new Error(
                "No se encontró en PermisosTI un registro cuyo campo " +
                "'UsuarioCorreo' coincida con la cuenta de Microsoft. " +
                "Revisa el correo registrado en SharePoint."
            );
        }

        const grupo = String(
            obtenerCampo(registro, ["Grupo"]) ?? ""
        ).trim().toLowerCase();

        const correoRegistrado = normalizarCorreo(
            obtenerCampo(registro, [
                "UsuarioCorreo",
                "Usuario Correo",
                "Correo",
                "Email",
                "Mail"
            ])
        );

        if (grupo === "ti") {
            return {
                correo: correoRegistrado,
                grupo: "TI",
                modulos: [...MODULOS]
            };
        }

        if (grupo !== "normal") {
            throw new Error(
                "El registro de PermisosTI tiene un Grupo vacío o inválido. " +
                "Debe ser TI o Normal."
            );
        }

        const modulos = MODULOS.filter(function (modulo) {
            const valor = obtenerCampo(registro, [modulo]);
            return esVerdadero(valor);
        });

        console.info(
            "[ALFERZA ACCESO] Módulos autorizados:",
            modulos
        );

        return {
            correo: correoRegistrado,
            grupo: "Normal",
            modulos: modulos
        };
    }

    function obtenerModuloActual() {
        const archivo = window.location.pathname
            .toLowerCase()
            .split("/")
            .pop();

        if (!archivo || archivo === "index.html") {
            return "Oficina";
        }

        if (archivo === "permisos.html") {
            return "Permisos";
        }

        return PAGINAS[archivo] || null;
    }

    function filtrarMenu(permisos) {
        document.querySelectorAll("a[href]").forEach(function (enlace) {
            let destino;

            try {
                destino = new URL(enlace.getAttribute("href"), location.href);
            } catch (_) {
                return;
            }

            if (destino.origin !== location.origin) return;

            const archivo = destino.pathname.toLowerCase().split("/").pop();

            let modulo = PAGINAS[archivo] || null;

            if (archivo === "" || archivo === "index.html") {
                modulo = "Oficina";
            }

            if (archivo === "permisos.html") {
                modulo = "Permisos";
            }

            if (!modulo) return;

            const permitido = modulo === "Permisos"
                ? permisos.grupo === "TI"
                : permisos.modulos.includes(modulo);

            if (!permitido) {
                const contenedor = enlace.closest("li") || enlace;
                contenedor.style.display = "none";
                enlace.setAttribute("aria-hidden", "true");
                enlace.setAttribute("tabindex", "-1");
            }
        });
    }

    async function validarAcceso() {
        const archivo = window.location.pathname
            .toLowerCase()
            .split("/")
            .pop();

        if (PAGINAS_PUBLICAS.includes(archivo)) {
            return { autorizado: true, publico: true };
        }

        if (sessionStorage.getItem("alferza_login") !== "true") {
            sessionStorage.setItem(
                "alferza_return_url",
                window.location.pathname +
                window.location.search +
                window.location.hash
            );

            window.location.replace(CONFIG.basePath + "login.html");

            return { autorizado: false };
        }

        if (typeof msal === "undefined") {
            throw new Error("No se cargó la librería MSAL.");
        }

        instanciaMSAL = new msal.PublicClientApplication({
            auth: {
                clientId: CONFIG.clientId,
                authority:
                    "https://login.microsoftonline.com/" + CONFIG.tenantId,
                redirectUri: CONFIG.redirectUri
            },
            cache: {
                cacheLocation: "sessionStorage",
                storeAuthStateInCookie: false
            }
        });

        await instanciaMSAL.initialize();

        const respuestaRedirect =
            await instanciaMSAL.handleRedirectPromise();

        if (
            respuestaRedirect &&
            respuestaRedirect.account
        ) {
            instanciaMSAL.setActiveAccount(respuestaRedirect.account);
        }

        const cuentas = instanciaMSAL.getAllAccounts();

        if (!cuentas.length) {
            sessionStorage.removeItem("alferza_login");
            window.location.replace(CONFIG.basePath + "login.html");

            return { autorizado: false };
        }

        const cuentaActiva = instanciaMSAL.getActiveAccount();

        const cuentaGuardada = normalizarCorreo(
            sessionStorage.getItem("alferza_account_username")
        );

        cuentaActual =
            (cuentaActiva &&
                cuentas.find(c =>
                    c.homeAccountId === cuentaActiva.homeAccountId
                )) ||
            (cuentaGuardada &&
                cuentas.find(c =>
                    normalizarCorreo(c.username) === cuentaGuardada
                )) ||
            (cuentas.length === 1 ? cuentas[0] : null);

        if (!cuentaActual) {
            throw new Error(
                "Hay varias cuentas de Microsoft en la sesión y no se pudo " +
                "determinar cuál corresponde al inicio de sesión. " +
                "Cierra sesión y vuelve a entrar con la cuenta autorizada."
            );
        }

        const tenantCuenta = String(
            cuentaActual.tenantId ||
            cuentaActual.idTokenClaims?.tid ||
            ""
        ).toLowerCase();

        if (tenantCuenta && tenantCuenta !== CONFIG.tenantId.toLowerCase()) {
            throw new Error(
                "La cuenta seleccionada pertenece a otro tenant de Microsoft."
            );
        }

        instanciaMSAL.setActiveAccount(cuentaActual);

        const respuestaToken = await instanciaMSAL.acquireTokenSilent({
            account: cuentaActual,
            scopes: ["User.Read", "Sites.Read.All"]
        });

        tokenGraph = respuestaToken.accessToken;

        const usuario = await graph("/me?$select=mail,userPrincipalName");

        const correos = [
            usuario.mail,
            usuario.userPrincipalName,
            cuentaActual.username
        ].map(normalizarCorreo).filter(Boolean);

        const correosUnicos = [...new Set(correos)];

        if (!correosUnicos.length) {
            throw new Error(
                "Microsoft Graph no devolvió un correo identificable para la cuenta."
            );
        }

        console.info("[ALFERZA ACCESO] Cuenta validada:", {
            correoGraph: usuario.mail || "",
            userPrincipalName: usuario.userPrincipalName || "",
            cuentaMSAL: cuentaActual.username || ""
        });

        // Se vuelve a consultar SharePoint para no conservar permisos antiguos.
        sessionStorage.removeItem("alferza_permisos");

        const correoAdmin = correosUnicos.some(correo =>
            CONFIG.adminEmails
                .map(normalizarCorreo)
                .includes(correo)
        );

        let permisos;

        if (correoAdmin) {
            permisos = {
                correo: correosUnicos[0],
                grupo: "TI",
                modulos: [...MODULOS]
            };
        } else {
            permisos = await obtenerPermisos(correosUnicos);
        }

        const moduloActual = obtenerModuloActual();

        if (
            moduloActual === "Permisos" &&
            permisos.grupo !== "TI"
        ) {
            throw new Error(
                "La administración de permisos está reservada al grupo TI."
            );
        }

        if (
            moduloActual &&
            moduloActual !== "Permisos" &&
            !permisos.modulos.includes(moduloActual)
        ) {
            throw new Error(
                "La cuenta fue identificada como '" + permisos.correo +
                "', pero no tiene habilitado el módulo '" +
                moduloActual + "' en PermisosTI."
            );
        }

        window.alferzaPermisos = permisos;
        window.alferzaMSAL = instanciaMSAL;

        sessionStorage.setItem(
            "alferza_account_username",
            cuentaActual.username || ""
        );

        if (document.readyState === "loading") {
            await new Promise(function (resolve) {
                document.addEventListener(
                    "DOMContentLoaded",
                    resolve,
                    { once: true }
                );
            });
        }

        filtrarMenu(permisos);

        window.dispatchEvent(
            new CustomEvent("alferza:access-granted", {
                detail: permisos
            })
        );

        console.info("[ALFERZA ACCESO] Acceso autorizado:", {
            correo: permisos.correo,
            grupo: permisos.grupo,
            modulos: permisos.modulos
        });

        return {
            autorizado: true,
            permisos: permisos
        };
    }

    window.alferzaAccessReady = validarAcceso().catch(function (error) {
        console.error("[ALFERZA ACCESO] Error de validación:", error);

        mostrarDenegacion(
            "No se pudo autorizar el acceso. Revisa el detalle para identificar la causa.",
            false,
            error && error.message
                ? error.message
                : "Error desconocido durante la validación."
        );

        return {
            autorizado: false,
            error: error && error.message
                ? error.message
                : "Error desconocido"
        };
    });

})();
