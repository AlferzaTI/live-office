
import json
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path

import msal
import msal_extensions
import requests


# =========================================================
# CONFIGURACION
# =========================================================

def obtener_carpeta():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CARPETA = obtener_carpeta()
ARCHIVO_CONFIG = CARPETA / "config.json"
ARCHIVO_CACHE = CARPETA / "cache_msal.bin"

if not ARCHIVO_CONFIG.exists():
    raise FileNotFoundError(
        f"No se encontró config.json en: {CARPETA}"
    )

with open(ARCHIVO_CONFIG, "r", encoding="utf-8") as archivo:
    CONFIG = json.load(archivo)

CLIENT_ID = CONFIG["client_id"]
TENANT_ID = CONFIG["tenant_id"]
HOST = CONFIG["sharepoint_host"]
SITE_PATH = CONFIG["sharepoint_site_path"]
LIST_NAME = CONFIG["list_name"]
INTERVALO = max(15, int(CONFIG.get("refresh_seconds", 30)))

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Sites.Read.All"]

COLOR_FONDO = "#00205c"
COLOR_TEXTO = "#FFFFFF"

sesion_http = requests.Session()

token_actual = None
sitio_id = None
lista_id = None
cerrando = False

# Evita abrir varias ventanas de autenticación en la misma ejecución.
intento_interactivo_realizado = False

bloqueo_auth = threading.RLock()


# =========================================================
# CACHE PERSISTENTE DE MICROSOFT
# =========================================================

persistencia = msal_extensions.FilePersistenceWithDataProtection(
    str(ARCHIVO_CACHE)
)

cache_msal = msal_extensions.PersistedTokenCache(persistencia)

app_msal = msal.PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    token_cache=cache_msal
)


# =========================================================
# AUTENTICACION MICROSOFT
# =========================================================

def obtener_token(
    permitir_interactivo=False,
    forzar_renovacion=False
):
    global token_actual, intento_interactivo_realizado

    with bloqueo_auth:

        # 1. Buscar una sesión guardada.
        cuentas = app_msal.get_accounts()

        if cuentas:
            resultado = app_msal.acquire_token_silent(
                SCOPES,
                account=cuentas[0],
                force_refresh=forzar_renovacion
            )

            if resultado and resultado.get("access_token"):
                token_actual = resultado["access_token"]
                return token_actual

            if resultado:
                print(
                    "Renovación silenciosa:",
                    resultado.get(
                        "error_description",
                        resultado.get("error", "Sin detalles")
                    ),
                    flush=True
                )

        # 2. Si no hay sesión y no se permite iniciar sesión,
        # devolver un error sin abrir otra ventana.
        if not permitir_interactivo:
            raise RuntimeError(
                "No se pudo renovar la sesión de Microsoft. "
                "Cierra y vuelve a abrir el contador para iniciar "
                "sesión nuevamente."
            )

        if intento_interactivo_realizado:
            raise RuntimeError(
                "Ya se intentó iniciar sesión en esta ejecución. "
                "Cierra y vuelve a abrir el contador."
            )

        intento_interactivo_realizado = True

        # 3. Inicio de sesión normal de Microsoft.
        print(
            "Abriendo inicio de sesión de Microsoft...",
            flush=True
        )

        try:
            resultado = app_msal.acquire_token_interactive(
                scopes=SCOPES,
                prompt="select_account"
            )
        except Exception as error:
            raise RuntimeError(
                f"No se pudo abrir el inicio de sesión: {error}"
            )

        if not resultado or not resultado.get("access_token"):
            detalle = (
                resultado.get(
                    "error_description",
                    resultado.get("error", "Error desconocido")
                )
                if resultado else
                "Microsoft no devolvió un token."
            )

            raise RuntimeError(
                f"Falló el inicio de sesión: {detalle}"
            )

        token_actual = resultado["access_token"]

        print(
            "Inicio de sesión completado correctamente.",
            flush=True
        )

        return token_actual


# =========================================================
# MICROSOFT GRAPH
# =========================================================

def interpretar_error(respuesta):
    try:
        datos = respuesta.json()
        error = datos.get("error", {})

        codigo = error.get("code", "GraphError")
        mensaje = error.get("message", respuesta.text)

    except (ValueError, AttributeError):
        codigo = f"HTTP {respuesta.status_code}"
        mensaje = respuesta.text

    return f"{codigo} ({respuesta.status_code}): {mensaje}"


def graph_get(url):
    global token_actual

    # La primera conexión puede abrir el inicio de sesión.
    permitir_interactivo = not intento_interactivo_realizado

    token = obtener_token(
        permitir_interactivo=permitir_interactivo
    )

    respuesta = sesion_http.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=30
    )

    # Si Microsoft rechaza el token, intentamos renovarlo.
    if respuesta.status_code == 401:
        token_actual = None

        token = obtener_token(
            permitir_interactivo=False,
            forzar_renovacion=True
        )

        respuesta = sesion_http.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            timeout=30
        )

    if not respuesta.ok:
        raise RuntimeError(interpretar_error(respuesta))

    return respuesta.json()


# =========================================================
# CONSULTA DE SHAREPOINT
# =========================================================

def obtener_ids():
    global sitio_id, lista_id

    if sitio_id and lista_id:
        return sitio_id, lista_id

    url_sitio = (
        "https://graph.microsoft.com/v1.0/sites/"
        f"{HOST}:{SITE_PATH}"
    )

    sitio = graph_get(url_sitio)
    sitio_id = sitio["id"]

    url_listas = (
        "https://graph.microsoft.com/v1.0/sites/"
        f"{sitio_id}/lists"
        "?$select=id,name,displayName&$top=200"
    )

    resultado_listas = graph_get(url_listas)
    listas = resultado_listas.get("value", [])

    # Seguir paginación por si hay muchas listas.
    siguiente = resultado_listas.get("@odata.nextLink")

    while siguiente:
        pagina = graph_get(siguiente)
        listas.extend(pagina.get("value", []))
        siguiente = pagina.get("@odata.nextLink")

    objetivo = LIST_NAME.strip().casefold()

    lista = next(
        (
            item for item in listas
            if str(item.get("displayName", "")).strip().casefold()
            == objetivo
            or str(item.get("name", "")).strip().casefold()
            == objetivo
        ),
        None
    )

    if not lista:
        raise RuntimeError(
            f"No se encontró la lista '{LIST_NAME}'. "
            "Comprueba el nombre de la lista."
        )

    lista_id = lista["id"]

    print(
        f"SharePoint conectado. Lista: "
        f"{lista.get('displayName', LIST_NAME)}",
        flush=True
    )

    return sitio_id, lista_id


def obtener_conteos():
    sitio, lista = obtener_ids()

    url = (
        "https://graph.microsoft.com/v1.0/sites/"
        f"{sitio}/lists/{lista}/items"
        "?$expand=fields($select=Estado)&$top=200"
    )

    pendientes = 0
    proceso = 0

    while url:
        resultado = graph_get(url)

        for item in resultado.get("value", []):
            fields = item.get("fields", {})

            estado = str(
                fields.get("Estado", "")
            ).strip().casefold()

            if estado == "pendiente":
                pendientes += 1

            elif estado == "en proceso":
                proceso += 1

        url = resultado.get("@odata.nextLink")

    return pendientes, proceso


# =========================================================
# VENTANA FLOTANTE
# =========================================================

root = tk.Tk()
root.title("ALFERZA - ConteoTickets")
root.geometry("220x185+30+120")
root.resizable(False, False)
root.configure(bg=COLOR_FONDO)
root.attributes("-topmost", True)

try:
    root.attributes("-alpha", 0.97)
except tk.TclError:
    pass

barra = tk.Frame(
    root,
    bg=COLOR_FONDO,
    height=38
)
barra.pack(fill="x")
barra.pack_propagate(False)

titulo = tk.Label(
    barra,
    text="ALFERZA | TICKETS TI",
    bg=COLOR_FONDO,
    fg=COLOR_TEXTO,
    font=("Segoe UI", 9, "bold")
)
titulo.pack(side="left", padx=9)

contenido = tk.Frame(
    root,
    bg=COLOR_FONDO
)
contenido.pack(
    fill="both",
    expand=True,
    padx=12,
    pady=(3, 8)
)

tk.Label(
    contenido,
    text="PENDIENTES",
    bg=COLOR_FONDO,
    fg="#B8D7FF",
    font=("Segoe UI", 9, "bold")
).pack(anchor="w")

etiqueta_pendientes = tk.Label(
    contenido,
    text="—",
    bg=COLOR_FONDO,
    fg=COLOR_TEXTO,
    font=("Segoe UI", 27, "bold")
)
etiqueta_pendientes.pack(anchor="w")

fila = tk.Frame(
    contenido,
    bg=COLOR_FONDO
)
fila.pack(fill="x")

tk.Label(
    fila,
    text="En proceso:",
    bg=COLOR_FONDO,
    fg="#DCE8FF",
    font=("Segoe UI", 9)
).pack(side="left")

etiqueta_proceso = tk.Label(
    fila,
    text="—",
    bg=COLOR_FONDO,
    fg="#60A5FA",
    font=("Segoe UI", 10, "bold")
)
etiqueta_proceso.pack(side="left", padx=5)

etiqueta_estado = tk.Label(
    contenido,
    text="Preparando conexión...",
    bg=COLOR_FONDO,
    fg="#CBD5E1",
    font=("Segoe UI", 8),
    wraplength=195,
    justify="left"
)
etiqueta_estado.pack(
    anchor="w",
    pady=(5, 0)
)


# =========================================================
# ARRASTRAR Y CERRAR
# =========================================================

posicion = {"x": 0, "y": 0}


def iniciar_arrastre(event):
    posicion["x"] = event.x_root - root.winfo_x()
    posicion["y"] = event.y_root - root.winfo_y()


def arrastrar(event):
    x = event.x_root - posicion["x"]
    y = event.y_root - posicion["y"]

    root.geometry(f"+{x}+{y}")


def al_cerrar():
    global cerrando

    cerrando = True
    root.destroy()


boton_cerrar = tk.Button(
    barra,
    text="×",
    command=al_cerrar,
    bg="#D92D20",
    fg="white",
    activebackground="#B42318",
    activeforeground="white",
    relief="flat",
    bd=0,
    font=("Segoe UI", 15, "bold"),
    cursor="hand2",
    width=2
)
boton_cerrar.pack(
    side="right",
    padx=4,
    pady=4
)

for elemento in (barra, titulo):
    elemento.bind("<Button-1>", iniciar_arrastre)
    elemento.bind("<B1-Motion>", arrastrar)


# =========================================================
# ACTUALIZACION DE INTERFAZ
# =========================================================

def actualizar_interfaz(pendientes, proceso):
    if cerrando:
        return

    etiqueta_pendientes.config(
        text=str(pendientes)
    )

    etiqueta_proceso.config(
        text=str(proceso)
    )

    etiqueta_estado.config(
        text=f"Actualizado {datetime.now():%H:%M:%S}",
        fg="#86EFAC"
    )


def mostrar_error(mensaje):
    if cerrando:
        return

    etiqueta_estado.config(
        text=mensaje[:180],
        fg="#FCA5A5"
    )


# =========================================================
# CICLO DE ACTUALIZACION
# =========================================================

def ciclo_actualizacion():
    while not cerrando:

        try:
            pendientes, proceso = obtener_conteos()

            if not cerrando:
                root.after(
                    0,
                    lambda p=pendientes, e=proceso:
                    actualizar_interfaz(p, e)
                )

        except Exception as error:
            detalle = str(error)

            print(
                f"[{datetime.now():%H:%M:%S}] "
                f"Error actualizando tickets: {detalle}",
                flush=True
            )

            if not cerrando:
                try:
                    root.after(
                        0,
                        lambda m=detalle:
                        mostrar_error(m)
                    )
                except tk.TclError:
                    return

        # Esperar el intervalo sin congelar la ventana.
        for _ in range(INTERVALO * 10):
            if cerrando:
                return

            threading.Event().wait(0.1)


# =========================================================
# INICIO
# =========================================================

root.protocol(
    "WM_DELETE_WINDOW",
    al_cerrar
)

threading.Thread(
    target=ciclo_actualizacion,
    daemon=True
).start()

root.mainloop()
