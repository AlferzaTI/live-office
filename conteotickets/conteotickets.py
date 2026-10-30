
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

# La carpeta img está al mismo nivel que conteotickets.
CARPETA_CUPHEAD = CARPETA.parent / "img" / "Cuphead" / "Cuphead"

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
                prompt="select_account",
                port=8400
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
# CONSULTA SHAREPOINT
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


# =========================================================
# CONVERTIR FECHA DE SHAREPOINT A FECHA LOCAL
# =========================================================

def obtener_fecha_local(valor_fecha):
    """
    Convierte una fecha ISO de SharePoint a la fecha local
    de Windows.

    Acepta fechas como:
    2026-10-29T14:30:00Z
    2026-10-29T14:30:00+00:00

    Si la fecha no contiene zona horaria, utiliza su fecha
    directamente.
    """

    if not valor_fecha:
        return None

    try:
        if isinstance(valor_fecha, datetime):
            fecha = valor_fecha
        else:
            texto = str(valor_fecha).strip()

            if not texto:
                return None

            # Compatibilidad con fechas UTC terminadas en Z.
            if texto.endswith("Z"):
                texto = texto[:-1] + "+00:00"

            fecha = datetime.fromisoformat(texto)

        # Si SharePoint devuelve una fecha con zona horaria,
        # la convertimos a la zona horaria local de Windows.
        if fecha.tzinfo is not None:
            fecha = fecha.astimezone()

        return fecha.date()

    except (ValueError, TypeError, OverflowError) as error:
        print(
            f"No se pudo interpretar la fecha '{valor_fecha}': {error}",
            flush=True
        )
        return None


# =========================================================
# OBTENER PENDIENTES Y TICKETS REGISTRADOS HOY
# =========================================================

def obtener_conteos():
    sitio, lista = obtener_ids()

    # Solicitamos el Estado y el campo Created de SharePoint.
    # createdDateTime también se recibe como propiedad del elemento.
    url = (
        "https://graph.microsoft.com/v1.0/sites/"
        f"{sitio}/lists/{lista}/items"
        "?$expand=fields($select=Estado,Created)&$top=200"
    )

    pendientes = 0
    registrados_hoy = 0

    # Fecha actual de la computadora.
    # La computadora debe tener configurada la zona horaria de Perú.
    hoy = datetime.now().astimezone().date()

    while url:
        resultado = graph_get(url)

        for item in resultado.get("value", []):

            fields = item.get("fields", {})

            # -------------------------------------------------
            # 1. CONTAR TICKETS PENDIENTES
            # -------------------------------------------------

            estado = str(
                fields.get("Estado", "")
            ).strip().casefold()

            if estado == "pendiente":
                pendientes += 1

            # -------------------------------------------------
            # 2. CONTAR TICKETS CREADOS HOY
            # -------------------------------------------------

            # Primero usamos createdDateTime, que es el metadato
            # de creación del elemento en Microsoft Graph.
            #
            # Si no está disponible, utilizamos Created,
            # el campo de creación integrado de SharePoint.

            fecha_creacion = (
                item.get("createdDateTime")
                or fields.get("Created")
            )

            fecha_local = obtener_fecha_local(fecha_creacion)

            if fecha_local == hoy:
                registrados_hoy += 1

        # Continuar leyendo todas las páginas de la lista.
        url = resultado.get("@odata.nextLink")

    print(
        f"Pendientes: {pendientes} | "
        f"Registrados hoy: {registrados_hoy}",
        flush=True
    )

    return pendientes, registrados_hoy


# =========================================================
# DEFINICION DE ANIMACIONES CUPHEAD
# =========================================================

ANIMACIONES_CUPHEAD = {
    "idle": {
        "prefijo": "cuphead_intro_no_cookie_",
        "inicio": 1,
        "fin": 53,
        "velocidad": 80
    },
    "click": {
        "prefijo": "cuphead_intro_b_",
        "inicio": 1,
        "fin": 44,
        "velocidad": 80
    },
    "reporte": {
        "prefijo": "ch_intro_scared_portal_",
        "inicio": 1,
        "fin": 55,
        "velocidad": 80
    }
}

TAMANO_CUPHEAD = (95, 95)

frames_animaciones = {}
animacion_actual = None
indice_frame = 0
id_animacion = None
version_animacion = 0
reporte_en_curso = False


def cargar_animacion(nombre, configuracion):
    frames = []

    prefijo = configuracion["prefijo"]
    inicio = configuracion["inicio"]
    fin = configuracion["fin"]

    encontrados = 0

    for i in range(inicio, fin + 1):

        numero = str(i).zfill(4)

        archivo = CARPETA_CUPHEAD / f"{prefijo}{numero}.png"

        if not archivo.exists():
            print(
                f"No se encontró el frame: {archivo}",
                flush=True
            )
            continue

        try:
            with Image.open(archivo) as imagen_original:

                imagen = imagen_original.convert("RGBA")

                imagen.thumbnail(
                    TAMANO_CUPHEAD,
                    Image.Resampling.LANCZOS
                )

                frames.append(ImageTk.PhotoImage(imagen))

                encontrados += 1

        except Exception as error:
            print(
                f"Error cargando {archivo.name}: {error}",
                flush=True
            )

    print(
        f"Animación {nombre}: {encontrados} de "
        f"{fin - inicio + 1} frames cargados.",
        flush=True
    )

    return frames


def cargar_todas_las_animaciones():
    for nombre, configuracion in ANIMACIONES_CUPHEAD.items():
        frames_animaciones[nombre] = cargar_animacion(
            nombre,
            configuracion
        )


# =========================================================
# VENTANA PRINCIPAL
# =========================================================

root = tk.Tk()
root.title("ALFERZA | ConteoTickets")
root.overrideredirect(True)
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
    global cerrando, version_animacion

    if cerrando:
        return

    cerrando = True
    version_animacion += 1
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


# =========================================================
# ZONA DE ANIMACION CUPHEAD
# =========================================================

zona_cuphead = tk.Frame(
    panel_principal,
    bg=COLOR_PANEL,
    width=115
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


# Cargar los tres grupos de frames.
cargar_todas_las_animaciones()


# =========================================================
# CONTROLADOR DE ANIMACIONES
# =========================================================

def mostrar_frame(frames, indice):
    if not frames or cerrando:
        return

    indice_seguro = indice % len(frames)

    etiqueta_cuphead.config(
        image=frames[indice_seguro],
        text=""
    )


def cancelar_animacion_programada():
    global id_animacion

    if id_animacion is not None:
        try:
            root.after_cancel(id_animacion)
        except tk.TclError:
            pass

        id_animacion = None


def iniciar_animacion(nombre, repetir=True, al_terminar=None):
    global animacion_actual
    global indice_frame
    global version_animacion
    global id_animacion

    if cerrando:
        return

    frames = frames_animaciones.get(nombre, [])

    cancelar_animacion_programada()

    version_animacion += 1
    version_actual = version_animacion

    animacion_actual = nombre
    indice_frame = 0

    if not frames:
        if al_terminar:
            al_terminar()
        elif nombre != "idle":
            iniciar_animacion("idle", repetir=True)
        return

    configuracion = ANIMACIONES_CUPHEAD[nombre]
    velocidad = configuracion["velocidad"]

    def avanzar():
        global indice_frame, id_animacion

        if cerrando or version_actual != version_animacion:
            return

        if indice_frame >= len(frames):

            if repetir:
                indice_frame = 0

            else:
                id_animacion = None

                if al_terminar:
                    al_terminar()
                elif nombre != "idle":
                    iniciar_animacion("idle", repetir=True)

                return

        mostrar_frame(frames, indice_frame)
        indice_frame += 1

        id_animacion = root.after(
            velocidad,
            avanzar
        )

    avanzar()


def animar_cuphead():
    iniciar_animacion("idle", repetir=True)


# =========================================================
# PANEL DE TICKETS REGISTRADOS HOY
# =========================================================

panel_hoy = tk.Frame(
    contenido,
    bg=COLOR_PANEL_CLARO
)

panel_hoy.pack(
    fill="x",
    pady=(0, 9)
)

fila_hoy = tk.Frame(
    panel_hoy,
    bg=COLOR_PANEL_CLARO
)

fila_hoy.pack(
    fill="x",
    padx=13,
    pady=8
)

tk.Label(
    fila_hoy,
    text="TICKETS REGISTRADOS HOY",
    bg=COLOR_PANEL_CLARO,
    fg=COLOR_SECUNDARIO,
    font=("Segoe UI", 8, "bold")
).pack(side="left")

etiqueta_hoy = tk.Label(
    fila_hoy,
    text="—",
    bg=COLOR_PANEL_CLARO,
    fg=COLOR_VERDE,
    font=("Segoe UI", 15, "bold")
)

etiqueta_hoy.pack(side="right")


# =========================================================
# BOTON ABRIR REPORTE
# =========================================================

def abrir_reporte_en_navegador():
    global reporte_en_curso

    if cerrando:
        return

    try:
        resultado = webbrowser.open_new_tab(URL_REPORTE)

        if resultado:
            etiqueta_estado.config(
                text="Reporte abierto en el navegador",
                fg=COLOR_VERDE
            )
        else:
            etiqueta_estado.config(
                text="Se solicitó abrir el reporte",
                fg=COLOR_VERDE
            )

    except Exception as error:
        etiqueta_estado.config(
            text=f"No se pudo abrir el reporte: {error}",
            fg=COLOR_ROJO
        )

    reporte_en_curso = False

    if not cerrando:
        iniciar_animacion("idle", repetir=True)


def abrir_reporte():
    global reporte_en_curso

    if cerrando or reporte_en_curso:
        return

    reporte_en_curso = True

    etiqueta_estado.config(
        text="Abriendo reporte de tickets...",
        fg=COLOR_AZUL_CLARO
    )

    # Primero reproduce la animación especial.
    # Al terminar, abre el reporte en el navegador.
    iniciar_animacion(
        "reporte",
        repetir=False,
        al_terminar=abrir_reporte_en_navegador
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
# CLIC GENERAL EN LA VENTANA
# =========================================================

def pertenece_a(widget, elemento):
    actual = widget

    while actual is not None:
        if actual == elemento:
            return True

        actual = getattr(actual, "master", None)

    return False


def manejar_clic_ventana(event):
    if cerrando or reporte_en_curso:
        return

    widget = event.widget

    # No activar la animación general al pulsar botones.
    if pertenece_a(widget, boton_reporte):
        return

    if pertenece_a(widget, boton_cerrar):
        return

    # Un clic normal reproduce la animación de reacción.
    # Al terminar vuelve a la animación predeterminada.
    iniciar_animacion(
        "click",
        repetir=False
    )


root.bind_all(
    "<Button-1>",
    manejar_clic_ventana,
    add="+"
)


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

def actualizar_interfaz(pendientes, registrados_hoy):

    if cerrando:
        return

    etiqueta_pendientes.config(
        text=str(pendientes)
    )

    etiqueta_hoy.config(
        text=str(registrados_hoy)
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
# CICLO DE ACTUALIZACION
# =========================================================

def ciclo_actualizacion():

    while not evento_cierre.is_set():

        try:
            pendientes, registrados_hoy = obtener_conteos()

            if not evento_cierre.is_set():
                root.after(
                    0,
                    lambda p=pendientes, h=registrados_hoy:
                    actualizar_interfaz(p, h)
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

# Inicia con la animación predeterminada.
animar_cuphead()

threading.Thread(
    target=ciclo_actualizacion,
    daemon=True
).start()

root.mainloop()
