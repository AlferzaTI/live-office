
import json
import os
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from urllib.parse import quote

import msal
import requests


# =========================================================
# CONFIGURACION
# =========================================================

def obtener_carpeta():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CARPETA = obtener_carpeta()

with open(CARPETA / "config.json", "r", encoding="utf-8") as archivo:
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
COLOR_AZUL = "#0A84FF"
COLOR_TEXTO = "#FFFFFF"

sesion_http = requests.Session()
token_actual = None
sitio_id = None
lista_id = None
cerrando = False


# =========================================================
# AUTENTICACION MICROSOFT
# =========================================================

app_msal = msal.PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)


def obtener_token():
    global token_actual

    cuentas = app_msal.get_accounts()

    if cuentas:
        resultado = app_msal.acquire_token_silent(
            SCOPES,
            account=cuentas[0]
        )

        if resultado and resultado.get("access_token"):
            token_actual = resultado["access_token"]
            return token_actual

    # Autenticacion interactiva. Abre el navegador para iniciar sesion.
    resultado = app_msal.acquire_token_interactive(
        scopes=SCOPES,
        prompt="select_account"
    )

    if not resultado or not resultado.get("access_token"):
        raise RuntimeError(
            resultado.get(
                "error_description",
                "No se pudo autenticar con Microsoft."
            ) if resultado else "Microsoft no devolvio un token."
        )

    token_actual = resultado["access_token"]
    return token_actual


def graph_get(url):
    token = obtener_token()

    respuesta = sesion_http.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=25
    )

    if respuesta.status_code == 401:
        global token_actual
        token_actual = None

        token = obtener_token()

        respuesta = sesion_http.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            timeout=25
        )

    respuesta.raise_for_status()
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
        f"{sitio_id}/lists?$select=id,name,displayName&$top=200"
    )

    listas = graph_get(url_listas).get("value", [])

    objetivo = LIST_NAME.casefold()

    lista = next(
        (
            item for item in listas
            if str(item.get("displayName", "")).casefold() == objetivo
            or str(item.get("name", "")).casefold() == objetivo
        ),
        None
    )

    if not lista:
        raise RuntimeError(
            f"No se encontro la lista {LIST_NAME}."
        )

    lista_id = lista["id"]
    return sitio_id, lista_id


def obtener_conteos():
    sitio, lista = obtener_ids()

    url = (
        "https://graph.microsoft.com/v1.0/sites/"
        f"{sitio}/lists/{lista}/items"
        "?$expand=fields($select=Estado)&$top=999"
    )

    pendientes = 0
    proceso = 0

    while url:
        resultado = graph_get(url)

        for item in resultado.get("value", []):
            fields = item.get("fields", {})
            estado = str(fields.get("Estado", "Pendiente")).strip().casefold()

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

barra = tk.Frame(root, bg=COLOR_FONDO, height=38)
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

boton_cerrar = tk.Button(
    barra,
    text="×",
    command=root.destroy,
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
boton_cerrar.pack(side="right", padx=4, pady=4)

contenido = tk.Frame(root, bg=COLOR_FONDO)
contenido.pack(fill="both", expand=True, padx=12, pady=(3, 8))

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

fila = tk.Frame(contenido, bg=COLOR_FONDO)
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
    text="Conectando...",
    bg=COLOR_FONDO,
    fg="#CBD5E1",
    font=("Segoe UI", 8)
)
etiqueta_estado.pack(anchor="w", pady=(5, 0))


# =========================================================
# ARRASTRAR LA VENTANA
# =========================================================

posicion = {"x": 0, "y": 0}


def iniciar_arrastre(event):
    posicion["x"] = event.x_root - root.winfo_x()
    posicion["y"] = event.y_root - root.winfo_y()


def arrastrar(event):
    x = event.x_root - posicion["x"]
    y = event.y_root - posicion["y"]
    root.geometry(f"+{x}+{y}")


for elemento in (barra, titulo):
    elemento.bind("<Button-1>", iniciar_arrastre)
    elemento.bind("<B1-Motion>", arrastrar)


# =========================================================
# ACTUALIZACION AUTOMATICA
# =========================================================

def actualizar_interfaz(pendientes, proceso):
    if cerrando:
        return

    etiqueta_pendientes.config(text=str(pendientes))
    etiqueta_proceso.config(text=str(proceso))
    etiqueta_estado.config(
        text=f"Actualizado {__import__('datetime').datetime.now().strftime('%H:%M:%S')}",
        fg="#86EFAC"
    )


def mostrar_error(mensaje):
    if cerrando:
        return

    etiqueta_estado.config(
        text=mensaje[:55],
        fg="#FCA5A5"
    )


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
            print("Error actualizando tickets:", error)

            if not cerrando:
                root.after(
                    0,
                    lambda: mostrar_error("Error de conexion / acceso")
                )

        # Esperar el intervalo sin bloquear la ventana.
        for _ in range(INTERVALO * 10):
            if cerrando:
                return
            threading.Event().wait(0.1)


def al_cerrar():
    global cerrando
    cerrando = True
    root.destroy()


root.protocol("WM_DELETE_WINDOW", al_cerrar)

threading.Thread(
    target=ciclo_actualizacion,
    daemon=True
).start()

root.mainloop()
