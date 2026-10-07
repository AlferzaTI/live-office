document.addEventListener("DOMContentLoaded", () => {
    console.log("Permisos cargado");

    const contenido = document.querySelector(".content");

    if (contenido) {
        contenido.style.display = "block";
        contenido.style.visibility = "visible";
        contenido.style.opacity = "1";
    }

    inicializarPermisos();
});


async function inicializarPermisos() {

    try {

        const cuenta = await obtenerCuentaActual();

        if (!cuenta) {
            mostrarMensaje("No se pudo obtener el usuario actual.");
            return;
        }

        console.log("Usuario:", cuenta.mail || cuenta.userPrincipalName);

        cargarInterfazPermisos();

        await cargarPermisos();

    } catch (error) {

        console.error("Error en permisos:", error);

        cargarInterfazPermisos();

        mostrarMensaje(
            "No se pudieron cargar los permisos. Revisa la consola."
        );
    }
}


function cargarInterfazPermisos() {

    const contenido = document.querySelector(".content");

    if (!contenido) return;

    contenido.innerHTML = `
        <div class="page-header">
            <div>
                <h1>Permisos</h1>
                <p>Administración de accesos al sistema</p>
            </div>
        </div>

        <div class="permisos-container">

            <div class="permisos-card">

                <h2>Asignar permisos</h2>

                <div class="form-group">
                    <label for="correoUsuario">
                        Correo corporativo
                    </label>

                    <div class="input-row">
                        <input
                            type="email"
                            id="correoUsuario"
                            placeholder="usuario@alferza.pe"
                        >

                        <button
                            type="button"
                            id="btnBuscarUsuario"
                        >
                            Buscar
                        </button>
                    </div>
                </div>

                <div
                    id="panelPermisos"
                    style="display:none;"
                >

                    <div class="usuario-info">
                        <strong id="nombreUsuario">
                            Usuario
                        </strong>

                        <span id="correoMostrado">
                            usuario@alferza.pe
                        </span>
                    </div>

                    <div class="form-group">

                        <label for="grupoUsuario">
                            Grupo
                        </label>

                        <select id="grupoUsuario">
                            <option value="Normal">
                                Normal
                            </option>

                            <option value="TI">
                                TI
                            </option>
                        </select>

                    </div>

                    <div class="modulos-permisos">

                        ${crearModulo("Oficina")}
                        ${crearModulo("Personal")}
                        ${crearModulo("Reservas")}
                        ${crearModulo("Salas")}
                        ${crearModulo("Comunicados")}
                        ${crearModulo("Seguridad")}
                        ${crearModulo("Infraestructura")}
                        ${crearModulo("Tickets")}
                        ${crearModulo("Configuracion")}

                    </div>

                    <button
                        type="button"
                        id="btnGuardarPermisos"
                    >
                        Guardar permisos
                    </button>

                </div>

            </div>

            <div class="permisos-card">

                <h2>Usuarios configurados</h2>

                <div id="tablaPermisos">
                    Cargando usuarios...
                </div>

            </div>

        </div>
    `;

    document
        .getElementById("btnBuscarUsuario")
        ?.addEventListener(
            "click",
            buscarUsuario
        );

    document
        .getElementById("grupoUsuario")
        ?.addEventListener(
            "change",
            cambiarGrupo
        );

    document
        .getElementById("btnGuardarPermisos")
        ?.addEventListener(
            "click",
            guardarPermisos
        );
}


function crearModulo(nombre) {

    return `
        <div class="modulo-item">

            <span>${nombre}</span>

            <label class="switch">

                <input
                    type="checkbox"
                    class="permiso-modulo"
                    data-modulo="${nombre}"
                >

                <span class="slider"></span>

            </label>

        </div>
    `;
}


async function obtenerCuentaActual() {

    if (
        typeof msalInstance === "undefined" &&
        typeof window.msalInstance === "undefined"
    ) {
        console.error("MSAL no está disponible.");
        return null;
    }

    const instancia =
        window.msalInstance || msalInstance;

    let cuenta = instancia.getActiveAccount();

    if (!cuenta) {

        const cuentas =
            instancia.getAllAccounts();

        if (cuentas.length > 0) {
            cuenta = cuentas[0];
        }
    }

    if (!cuenta) {
        return null;
    }

    return {
        nombre:
            cuenta.name ||
            "",
        mail:
            cuenta.username ||
            cuenta.mail ||
            ""
    };
}


async function cargarPermisos() {

    const tabla =
        document.getElementById("tablaPermisos");

    if (!tabla) return;

    tabla.innerHTML = `
        <p>No hay usuarios configurados.</p>
    `;

    /*
     * Aquí conectas Microsoft Lists.
     * La interfaz ya queda visible aunque Graph falle.
     */
}


async function buscarUsuario() {

    const correo =
        document
            .getElementById("correoUsuario")
            ?.value
            .trim();

    if (!correo) {
        alert("Ingresa un correo corporativo.");
        return;
    }

    document
        .getElementById("correoMostrado")
        .textContent = correo;

    document
        .getElementById("nombreUsuario")
        .textContent = correo;

    document
        .getElementById("panelPermisos")
        .style.display = "block";

    document
        .getElementById("grupoUsuario")
        .value = "Normal";

    cambiarGrupo();
}


function cambiarGrupo() {

    const grupo =
        document
            .getElementById("grupoUsuario")
            ?.value;

    const controles =
        document.querySelectorAll(
            ".permiso-modulo"
        );

    controles.forEach(control => {

        if (grupo === "TI") {

            control.checked = true;
            control.disabled = true;

        } else {

            control.disabled = false;

        }

    });
}


async function guardarPermisos() {

    const correo =
        document
            .getElementById("correoUsuario")
            .value
            .trim();

    const grupo =
        document
            .getElementById("grupoUsuario")
            .value;

    if (!correo) {
        alert("Ingresa un correo.");
        return;
    }

    const permisos = {};

    document
        .querySelectorAll(".permiso-modulo")
        .forEach(control => {

            permisos[
                control.dataset.modulo
            ] = control.checked;

        });

    console.log({
        correo,
        grupo,
        permisos
    });

    alert("Permisos preparados correctamente.");
}


function mostrarMensaje(mensaje) {

    const panel =
        document.querySelector(".content");

    if (!panel) return;

    const aviso =
        document.createElement("div");

    aviso.className = "mensaje-error";
    aviso.textContent = mensaje;

    panel.prepend(aviso);
}