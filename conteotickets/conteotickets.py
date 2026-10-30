
import json
import sys
import threading
import tkinter as tk
import webbrowser
from datetime import datetime
from pathlib import Path

import msal
import msal_extensions
import requests
from PIL import Image, ImageTk


# =========================================================
# CONFIGURACION ORIGINAL
# =========================================================

def obtener_carpeta():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def obtener_recursos():
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


CARPETA = obtener_carpeta()
RECURSOS = obtener_recursos()

ARCHIVO_CONFIG = CARPETA / "config.json"
ARCHIVO_CACHE = CARPETA / "cache_msal.bin"

CARPETA_CUPHEAD = (
    RECURSOS / "img" / "Cuphead" / "Cuphead"
)

URL_REPORTE = (
    "https://alferzati.github.io/live-office/tickets.html"
)

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

COLOR_FONDO = "#07162E"
COLOR_PANEL = "#102746"
COLOR_PANEL_CLARO = "#16365C"
COLOR_AZUL = "#1685F8"
COLOR_AZUL_CLARO = "#63B3FF"
COLOR_TEXTO = "#FFFFFF"
COLOR_SECUNDARIO = "#A9BDD8"
COLOR_VERDE = "#36D399"
COLOR_ROJO = "#F87171"

sesion_http = requests.Session()

token_actual = None
sitio_id = None
lista_id = None
cerrando = False

intento_interactivo_realizado = False
bloqueo_auth = threading.RLock()
evento_cierre = threading.Event()


# =========================================================
# CACHE PERSISTENTE ORIGINAL
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
# AUTENTICACION ORIGINAL
# =========================================================

def obtener_token(
    permitir_interactivo=False,
    forzar_renovacion=False
):
    global token_actual, intento_interactivo_realizado

    with bloqueo_auth:

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
# MICROSOFT GRAPH ORIGINAL
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

    permitir_interactivo = not intento_interactivo_realizado

    token = obtener_token(
        permitir_interactivo=permitir_interactivo
    )

    respuesta = sesion_http.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=30
    )

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
# CONSULTA SHAREPOINT ORIGINAL
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
# CARGA DE LOS 44 FRAMES DE CUPHEAD
# =========================================================

def cargar_frames():
    frames = []

    for i in range(1, 45):

        numero = str(i).zfill(4)

        archivo = (
            CARPETA_CUPHEAD
            / f"cuphead_intro_b_{numero}.png"
        )

        if not archivo.exists():
            print(
                f"No se encontró el frame: {archivo}",
                flush=True
            )
            continue

        try:
            imagen = Image.open(archivo).convert("RGBA")

            imagen.thumbnail(
                (82, 82),
                Image.Resampling.LANCZOS
            )

            frames.append(ImageTk.PhotoImage(imagen))

        except Exception as error:
            print(
                f"Error cargando frame {i}: {error}",
                flush=True
            )

    print(
        f"Cuphead: {len(frames)} de 44 frames cargados.",
        flush=True
    )

    return frames


# =========================================================
# VENTANA PRINCIPAL
# =========================================================

root = tk.Tk()
root.title("ALFERZA | ConteoTickets")
root.geometry("365x280+30+120")
root.resizable(False, False)
root.configure(bg=COLOR_FONDO)
root.attributes("-topmost", True)

try:
    root.attributes("-alpha", 0.98)
except tk.TclError:
    pass


# =========================================================
# CABECERA
# =========================================================

barra = tk.Frame(
    root,
    bg=COLOR_FONDO,
    height=43
)

barra.pack(fill="x")
barra.pack_propagate(False)

titulo = tk.Label(
    barra,
    text="ALFERZA",
    bg=COLOR_FONDO,
    fg=COLOR_TEXTO,
    font=("Segoe UI", 12, "bold")
)

titulo.pack(
    side="left",
    padx=(14, 5),
    pady=8
)

subtitulo = tk.Label(
    barra,
    text="| LIVE OFFICE",
    bg=COLOR_FONDO,
    fg=COLOR_AZUL_CLARO,
    font=("Segoe UI", 9, "bold")
)

subtitulo.pack(
    side="left",
    pady=10
)

indicador_live = tk.Label(
    barra,
    text="● LIVE",
    bg=COLOR_FONDO,
    fg=COLOR_VERDE,
    font=("Segoe UI", 8, "bold")
)

indicador_live.pack(
    side="right",
    padx=(0, 7)
)


def al_cerrar():
    global cerrando

    if cerrando:
        return

    cerrando = True
    evento_cierre.set()

    try:
        root.destroy()
    except tk.TclError:
        pass


boton_cerrar = tk.Button(
    barra,
    text="×",
    command=al_cerrar,
    bg=COLOR_FONDO,
    fg=COLOR_SECUNDARIO,
    activebackground="#B42318",
    activeforeground=COLOR_TEXTO,
    relief="flat",
    bd=0,
    font=("Segoe UI", 15, "bold"),
    cursor="hand2",
    width=2
)

boton_cerrar.pack(
    side="right",
    padx=5
)


# =========================================================
# CONTENIDO
# =========================================================

contenido = tk.Frame(
    root,
    bg=COLOR_FONDO
)

contenido.pack(
    fill="both",
    expand=True,
    padx=12,
    pady=(2, 8)
)


# =========================================================
# PANEL DE PENDIENTES + CUPHEAD
# =========================================================

panel_principal = tk.Frame(
    contenido,
    bg=COLOR_PANEL,
    height=112
)

panel_principal.pack(
    fill="x",
    pady=(0, 9)
)

panel_principal.pack_propagate(False)

# Zona izquierda: contador

zona_pendientes = tk.Frame(
    panel_principal,
    bg=COLOR_PANEL
)

zona_pendientes.pack(
    side="left",
    fill="both",
    expand=True,
    padx=(13, 0),
    pady=9
)

tk.Label(
    zona_pendientes,
    text="TICKETS PENDIENTES",
    bg=COLOR_PANEL,
    fg=COLOR_SECUNDARIO,
    font=("Segoe UI", 8, "bold")
).pack(anchor="w")

etiqueta_pendientes = tk.Label(
    zona_pendientes,
    text="—",
    bg=COLOR_PANEL,
    fg=COLOR_TEXTO,
    font=("Segoe UI", 35, "bold")
)

etiqueta_pendientes.pack(
    anchor="w",
    pady=(0, 0)
)

tk.Label(
    zona_pendientes,
    text="Requieren atención",
    bg=COLOR_PANEL,
    fg=COLOR_AZUL_CLARO,
    font=("Segoe UI", 8)
).pack(anchor="w")


# Zona derecha: animación Cuphead

zona_cuphead = tk.Frame(
    panel_principal,
    bg=COLOR_PANEL,
    width=105
)

zona_cuphead.pack(
    side="right",
    fill="y",
    padx=(0, 8),
    pady=5
)

zona_cuphead.pack_propagate(False)

etiqueta_cuphead = tk.Label(
    zona_cuphead,
    bg=COLOR_PANEL,
    bd=0
)

etiqueta_cuphead.pack(
    expand=True
)

frames_cuphead = cargar_frames()
indice_frame = 0


def animar_cuphead():
    global indice_frame

    if cerrando:
        return

    if frames_cuphead:

        etiqueta_cuphead.config(
            image=frames_cuphead[indice_frame]
        )

        indice_frame += 1

        if indice_frame >= len(frames_cuphead):
            indice_frame = 0

    else:
        etiqueta_cuphead.config(
            text="ALFERZA",
            fg=COLOR_AZUL_CLARO,
            font=("Segoe UI", 9, "bold")
        )

    root.after(80, animar_cuphead)


# =========================================================
# PANEL EN PROCESO
# =========================================================

panel_proceso = tk.Frame(
    contenido,
    bg=COLOR_PANEL_CLARO
)

panel_proceso.pack(
    fill="x",
    pady=(0, 9)
)

fila_proceso = tk.Frame(
    panel_proceso,
    bg=COLOR_PANEL_CLARO
)

fila_proceso.pack(
    fill="x",
    padx=13,
    pady=8
)

tk.Label(
    fila_proceso,
    text="TICKETS EN PROCESO",
    bg=COLOR_PANEL_CLARO,
    fg=COLOR_SECUNDARIO,
    font=("Segoe UI", 8, "bold")
).pack(side="left")

etiqueta_proceso = tk.Label(
    fila_proceso,
    text="—",
    bg=COLOR_PANEL_CLARO,
    fg=COLOR_AZUL_CLARO,
    font=("Segoe UI", 15, "bold")
)

etiqueta_proceso.pack(side="right")


# =========================================================
# BOTON ABRIR REPORTE
# =========================================================

def abrir_reporte():

    try:
        webbrowser.open_new_tab(URL_REPORTE)

        etiqueta_estado.config(
            text="Reporte abierto en el navegador",
            fg=COLOR_VERDE
        )

    except Exception as error:

        etiqueta_estado.config(
            text=f"No se pudo abrir el reporte: {error}",
            fg=COLOR_ROJO
        )


boton_reporte = tk.Button(
    contenido,
    text="↗   ABRIR REPORTE DE TICKETS",
    command=abrir_reporte,
    bg=COLOR_AZUL,
    fg=COLOR_TEXTO,
    activebackground="#0867CC",
    activeforeground=COLOR_TEXTO,
    relief="flat",
    bd=0,
    font=("Segoe UI", 9, "bold"),
    cursor="hand2",
    pady=9
)

boton_reporte.pack(
    fill="x",
    pady=(0, 8)
)


def boton_entrar(event):
    boton_reporte.config(bg="#0867CC")


def boton_salir(event):
    boton_reporte.config(bg=COLOR_AZUL)


boton_reporte.bind("<Enter>", boton_entrar)
boton_reporte.bind("<Leave>", boton_salir)


# =========================================================
# ESTADO DE CONEXION
# =========================================================

etiqueta_estado = tk.Label(
    contenido,
    text="Preparando conexión con SharePoint...",
    bg=COLOR_FONDO,
    fg=COLOR_SECUNDARIO,
    font=("Segoe UI", 8),
    anchor="w"
)

etiqueta_estado.pack(fill="x")


# =========================================================
# ARRASTRAR VENTANA
# =========================================================

posicion = {"x": 0, "y": 0}


def iniciar_arrastre(event):
    posicion["x"] = event.x_root - root.winfo_x()
    posicion["y"] = event.y_root - root.winfo_y()


def arrastrar(event):
    x = event.x_root - posicion["x"]
    y = event.y_root - posicion["y"]

    root.geometry(f"+{x}+{y}")


# Solo la cabecera permite arrastrar.
for elemento in (
    barra,
    titulo,
    subtitulo,
    indicador_live
):
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
        text=f"● Actualizado: {datetime.now():%H:%M:%S}",
        fg=COLOR_VERDE
    )


def mostrar_error(mensaje):

    if cerrando:
        return

    etiqueta_estado.config(
        text=f"● {mensaje[:110]}",
        fg=COLOR_ROJO
    )


# =========================================================
# CICLO DE ACTUALIZACION ORIGINAL
# =========================================================

def ciclo_actualizacion():

    while not evento_cierre.is_set():

        try:
            pendientes, proceso = obtener_conteos()

            if not evento_cierre.is_set():
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

            if not evento_cierre.is_set():
                try:
                    root.after(
                        0,
                        lambda m=detalle:
                        mostrar_error(m)
                    )
                except tk.TclError:
                    return

        if evento_cierre.wait(INTERVALO):
            return


# =========================================================
# INICIO
# =========================================================

root.protocol(
    "WM_DELETE_WINDOW",
    al_cerrar
)

animar_cuphead()

threading.Thread(
    target=ciclo_actualizacion,
    daemon=True
).start()

root.mainloop()
