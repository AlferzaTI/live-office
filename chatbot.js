"use strict";

(() => {
  // =========================================================
  // ALFERZA CHATBOT v3
  // Microsoft Graph + SharePoint (PDF) + Microsoft Lists + Ollama
  // =========================================================

  // =========================================================
  // CONFIGURACIÓN
  // =========================================================

  const CFG = {
    clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
    tenantId: "dbab984f-4bb1-4b60-9dff-da59f54acdf1",

    redirectUri: (() => {
      const host = location.hostname;
      return host === "localhost" || host === "192.168.2.30"
        ? `${location.origin}/blank.html`
        : "https://alferzati.github.io/live-office/blank.html";
    })(),

    scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"],

    docs: {
      host: "alferzaholding.sharepoint.com",
      site: "/sites/TI",
      folder: "Procedimientos-TI"
    },

    lista: {
      host: "alferzaholding-my.sharepoint.com",
      site: "/personal/soporte1_alferza_pe",
      name: "Conversaciones-CHATBOT"
    },

    ollama: {
      url: "https://LT-0010.AD.ALFERZA/ollama/api/chat",
      model: "llama3:latest",   // Un modelo más pequeño (p. ej. llama3.2:3b) responde mucho más rápido.
      timeoutMs: 120000,         // Tiempo máximo total de una respuesta.
      stallMs: 60000,            // Tiempo máximo sin recibir texto nuevo.
      temperature: 0.2,
      numCtx: 3072,
      numPredict: 400,
      keepAlive: "60m"
    },

    pdfJs: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
    pdfWorker: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js",
    msalJs: "https://alcdn.msauth.net/browser/2.38.1/js/msal-browser.min.js",

    maxPaginasPdf: 300,
    concurrencia: 4,

    recuperacion: {
      topK: 3,           // Fragmentos enviados a la IA.
      maxChars: 2200,    // Límite de texto documental por consulta.
      chunkChars: 900,
      overlap: 150
    },

    historial: {
      max: 50,
      turnosIA: 4,       // Intercambios previos que la IA recuerda.
      recorteIA: 500     // Máximo de caracteres de cada mensaje previo enviado a la IA.
    },

    reintentoDocsMs: 60000
  };

  const GRAPH = "https://graph.microsoft.com/v1.0";

  const LS_KEYS = {
    history: "alferza_chat_history_v3",
    conv: "alferza_chat_conversation_v3"
  };

  const IDB_NAME = "alferza-chatbot";
  const IDB_STORE = "kv";

  // Sin tildes (el texto se normaliza antes de comparar).
  const STOP = new Set(
    (
      "a al algo algun alguna ante con como cual cuales cuando de del desde donde " +
      "e el ella ellas ellos en era es esa ese eso esta estas este esto estos fue " +
      "ha hacer hay la las le les lo los mas me mi mis mucho muy no o para pero por " +
      "porque puedo que quien quiero se sin sobre son su sus te tener tengo tu tus " +
      "un una uno unos y ya yo ayuda favor necesito quisiera dime hola gracias " +
      "buenas buenos dias tambien si"
    ).split(/\s+/)
  );

  // =========================================================
  // ESTADO
  // =========================================================

  const S = {
    msal: null,
    account: null,
    user: "",
    token: null,            // { valor, expira }
    sitioDocs: null,
    sitioLista: null,
    driveId: null,
    listaId: null,
    columnas: [],
    docs: [],
    indice: null,           // índice de búsqueda BM25
    historial: [],
    convId: "",
    ready: false,
    initPromise: null,
    ultimoIntentoInit: 0,
    busy: false,
    pdfJs: null
  };

  // =========================================================
  // UTILIDADES
  // =========================================================

  const $ = (s) => document.querySelector(s);
  const $$ = (s) => document.querySelectorAll(s);

  const log = (...a) => console.log("[ALFERZA]", ...a);
  const warn = (...a) => console.warn("[ALFERZA]", ...a);

  const norm = (t) =>
    String(t ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();

  const normalizar = (t) => norm(t).replace(/[^a-z0-9@._/-]+/g, " ").trim();
  const normCol = (t) => normalizar(t).replace(/\s/g, "");

  const uid = () =>
    crypto.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;

  const recortar = (t, max) => String(t ?? "").slice(0, max);

  const leerSesion = (k, def) => {
    try {
      return JSON.parse(sessionStorage.getItem(k)) ?? def;
    } catch {
      return def;
    }
  };

  const guardarSesion = (k, v) => {
    try {
      sessionStorage.setItem(k, JSON.stringify(v));
    } catch (e) {
      warn("No se pudo guardar en sessionStorage:", e);
    }
  };

  const el = (tag, clase, texto) => {
    const nodo = document.createElement(tag);
    if (clase) nodo.className = clase;
    if (texto != null) nodo.textContent = texto;
    return nodo;
  };

  // Ejecuta fn sobre los elementos con un máximo de `n` tareas simultáneas.
  async function enParalelo(items, n, fn) {
    let i = 0;
    const workers = Array.from({ length: Math.min(n, items.length) }, async () => {
      while (i < items.length) {
        const idx = i++;
        await fn(items[idx], idx);
      }
    });
    await Promise.all(workers);
  }

  // =========================================================
  // ALMACENAMIENTO PERSISTENTE (IndexedDB)
  // sessionStorage tiene poco espacio; aquí se guardan los PDF procesados.
  // =========================================================

  let dbPromise = null;

  function abrirDB() {
    if (!dbPromise) {
      dbPromise = new Promise((resolve, reject) => {
        const req = indexedDB.open(IDB_NAME, 1);
        req.onupgradeneeded = () => req.result.createObjectStore(IDB_STORE);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    return dbPromise;
  }

  async function idbGet(clave) {
    const db = await abrirDB();
    return new Promise((resolve, reject) => {
      const req = db.transaction(IDB_STORE).objectStore(IDB_STORE).get(clave);
      req.onsuccess = () => resolve(req.result ?? null);
      req.onerror = () => reject(req.error);
    });
  }

  async function idbSet(clave, valor) {
    const db = await abrirDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).put(valor, clave);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  // =========================================================
  // FORMATO DE RESPUESTAS (texto seguro, sin HTML)
  // =========================================================

  function textoConFormato(elemento, contenido) {
    const partes = String(contenido ?? "").split(/(\*\*[^*]+\*\*|`[^`]+`)/g);

    for (const parte of partes) {
      if (!parte) continue;

      if (parte.startsWith("**") && parte.endsWith("**") && parte.length > 4) {
        elemento.appendChild(el("strong", null, parte.slice(2, -2)));
      } else if (parte.startsWith("`") && parte.endsWith("`") && parte.length > 2) {
        elemento.appendChild(el("code", null, parte.slice(1, -1)));
      } else {
        elemento.appendChild(document.createTextNode(parte));
      }
    }
  }

  function formatearRespuesta(texto) {
    const cont = el("div", "response-formatted");
    let lista = null;

    const cerrarLista = () => {
      if (lista) {
        cont.appendChild(lista);
        lista = null;
      }
    };

    for (const original of String(texto || "").replace(/\r\n?/g, "\n").split("\n")) {
      const linea = original.trim();

      if (!linea) {
        cerrarLista();
        continue;
      }

      const titulo = linea.match(/^#{1,3}\s+(.+)$/);
      const numero = linea.match(/^\d+[.)]\s+(.+)$/);
      const vineta = linea.match(/^[-*•]\s+(.+)$/);

      if (titulo) {
        cerrarLista();
        const h = el("div", "response-heading");
        h.style.cssText = "font-weight:700;margin:10px 0 6px";
        textoConFormato(h, titulo[1]);
        cont.appendChild(h);
        continue;
      }

      if (numero || vineta) {
        const tipo = numero ? "ol" : "ul";

        if (!lista || lista.dataset.tipo !== tipo) {
          cerrarLista();
          lista = el(tipo);
          lista.dataset.tipo = tipo;
          lista.style.cssText = "margin:5px 0 10px;padding-left:24px";
        }

        const li = el("li");
        li.style.marginBottom = "5px";
        textoConFormato(li, (numero || vineta)[1]);
        lista.appendChild(li);
        continue;
      }

      cerrarLista();
      const p = el("p");
      p.style.cssText = "margin:0 0 9px;line-height:1.55";
      textoConFormato(p, linea);
      cont.appendChild(p);
    }

    cerrarLista();
    return cont;
  }

  // =========================================================
  // INTERFAZ
  // =========================================================

  const horaActual = () =>
    new Date().toLocaleTimeString("es-PE", { hour: "2-digit", minute: "2-digit" });

  function scrollAbajo() {
    const cont = $("#messages");
    if (cont) cont.scrollTop = cont.scrollHeight;
  }

  function crearMensaje(tipo) {
    const cont = $("#messages");
    const usuario = tipo === "user";

    const fila = el("div", `message-row ${usuario ? "user" : "bot"}`);
    const avatar = el("div", "message-avatar", usuario ? "TÚ" : "AI");
    const cuerpo = el("div", "message-content");
    const nombre = el("div", "message-name", usuario ? "Tú" : "Asistente ALFERZA");
    const burbuja = el("div", "message-bubble");
    const fuentes = el("div", "message-sources");
    fuentes.style.cssText = "font-size:12px;opacity:.8;margin-top:8px";
    fuentes.hidden = true;
    const hora = el("div", "message-time", horaActual());

    cuerpo.append(nombre, burbuja, fuentes, hora);
    fila.append(...(usuario ? [cuerpo, avatar] : [avatar, cuerpo]));

    cont?.appendChild(fila);
    scrollAbajo();

    return { burbuja, fuentes };
  }

  function pintarBot(msg, texto, fuentes = []) {
    msg.burbuja.replaceChildren(formatearRespuesta(texto));

    const unicas = [...new Set(fuentes)];
    msg.fuentes.textContent = unicas.length ? `Fuentes: ${unicas.join(", ")}` : "";
    msg.fuentes.hidden = unicas.length === 0;

    scrollAbajo();
  }

  // Actualiza la burbuja mientras llega la respuesta, limitado a un repintado por frame.
  function renderizadorDiferido(msg) {
    let pendiente = null;
    let frame = 0;

    const pintar = () => {
      frame = 0;
      msg.burbuja.replaceChildren(formatearRespuesta(pendiente));
      scrollAbajo();
    };

    return {
      actualizar(texto) {
        pendiente = texto;
        if (!frame) frame = requestAnimationFrame(pintar);
      },
      cancelar() {
        if (frame) cancelAnimationFrame(frame);
        frame = 0;
      }
    };
  }

  function estadoCarga(texto) {
    const cont = $("#messages");
    if (!cont) return;

    let aviso = $("#alferzaLoadingStatus");
    if (!aviso) {
      aviso = el("div");
      aviso.id = "alferzaLoadingStatus";
      aviso.style.cssText = "font-size:13px;opacity:.8;padding:8px 12px";
      cont.appendChild(aviso);
    }

    aviso.textContent = texto;
    scrollAbajo();
  }

  const ocultarCarga = () => $("#alferzaLoadingStatus")?.remove();

  function escribiendo(visible) {
    const indicador = $("#typingIndicator");
    if (indicador) {
      indicador.hidden = !visible;
      indicador.style.display = visible ? "flex" : "none";
    }
  }

  function actualizarContador() {
    const entrada = $("#messageInput");
    const contador = $("#characterCount");
    if (entrada && contador) contador.textContent = String(entrada.value.length);
  }

  function setControles() {
    const boton = $("#sendButton");
    if (boton) boton.disabled = S.busy;
    $$(".suggestion-button").forEach((b) => (b.disabled = S.busy));
  }

  function renderHistorial() {
    $("#messages")?.replaceChildren();

    if (!S.historial.length) {
      pintarBot(crearMensaje("bot"), "Hola. Soy el asistente de ALFERZA. ¿En qué te ayudo?");
      return;
    }

    for (const item of S.historial) {
      crearMensaje("user").burbuja.textContent = item.pregunta;
      pintarBot(crearMensaje("bot"), item.respuesta, item.fuentes || []);
    }
  }

  // =========================================================
  // SALUDOS Y CORTESÍA (respuesta instantánea, sin IA)
  // =========================================================

  function detectarSocial(pregunta) {
    const t = norm(pregunta)
      .replace(/[^a-z0-9 ]/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (!t || t.length > 40) return "";

    if (/^(hola+|hey|holi|buenas( tardes| noches| dias)?|buen(os)? dias?|saludos|que tal|hola buenas|hola que tal)( a todos| equipo| amigo)?( que tal| como estas| como va)?$/.test(t)) {
      return "saludo";
    }
    if (/^(muchas )?gracias( mil| totales)?$|^(ok|okay|vale|listo|perfecto|genial|excelente|entendido|claro|dale|ya|bien)( gracias)?$/.test(t)) {
      return "agradecimiento";
    }
    if (/^(chau|chao|adios|hasta luego|hasta pronto|nos vemos|bye|me voy)$/.test(t)) {
      return "despedida";
    }
    if (/^(como estas|como va|todo bien|que tal estas|como te va)$/.test(t)) {
      return "estado";
    }
    return "";
  }

  function respuestaSocial(tipo) {
    switch (tipo) {
      case "saludo":
        return "Hola. ¿En qué te ayudo?";
      case "agradecimiento":
        return "Perfecto. ¿Necesitas algo más?";
      case "despedida":
        return "Hasta luego.";
      case "estado":
        return "Bien, gracias. ¿En qué te ayudo?";
      default:
        return "¿En qué te ayudo?";
    }
  }

  // =========================================================
  // CARGA DE LIBRERÍAS
  // =========================================================

  async function cargarScript(src, global) {
    if (window[global]) return;

    await new Promise((resolve, reject) => {
      const script = el("script");
      script.src = src;
      script.onload = resolve;
      script.onerror = () => reject(new Error(`No se pudo cargar ${src}`));
      document.head.appendChild(script);
    });

    if (!window[global]) throw new Error(`${global} no está disponible.`);
  }

  async function cargarPdfJs() {
    if (S.pdfJs) return;

    await cargarScript(CFG.pdfJs, "pdfjsLib");
    S.pdfJs = window.pdfjsLib;
    S.pdfJs.GlobalWorkerOptions.workerSrc = CFG.pdfWorker;
  }

  // =========================================================
  // AUTENTICACIÓN MICROSOFT
  // =========================================================

  async function iniciarMSAL() {
    await cargarScript(CFG.msalJs, "msal");

    S.msal = new msal.PublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: `https://login.microsoftonline.com/${CFG.tenantId}`,
        redirectUri: CFG.redirectUri,
        navigateToLoginRequestUrl: false
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
      }
    });

    await S.msal.initialize();

    const redirect = await S.msal.handleRedirectPromise();
    S.account = redirect?.account || S.msal.getAllAccounts()[0] || null;

    if (!S.account) {
      const login = await S.msal.loginPopup({
        scopes: CFG.scopes,
        prompt: "select_account"
      });
      S.account = login.account;
    }

    if (!S.account) throw new Error("No se pudo identificar la cuenta Microsoft.");

    S.user = S.account.username || S.account.name || "";
    log("Sesión Microsoft lista:", S.user);
  }

  // Token en memoria: evita pedirlo a MSAL en cada llamada.
  async function obtenerToken() {
    if (S.token && S.token.expira - Date.now() > 60000) return S.token.valor;

    const pedir = (metodo) =>
      S.msal[metodo]({ scopes: CFG.scopes, account: S.account });

    let resultado;
    try {
      resultado = await pedir("acquireTokenSilent");
    } catch (e) {
      if (e.name !== "InteractionRequiredAuthError" && e.errorCode !== "interaction_required") {
        throw e;
      }
      resultado = await pedir("acquireTokenPopup");
    }

    S.token = {
      valor: resultado.accessToken,
      expira: new Date(resultado.expiresOn).getTime()
    };
    return S.token.valor;
  }

  // =========================================================
  // MICROSOFT GRAPH
  // =========================================================

  async function graph(url, opts = {}, intento = 0) {
    const accessToken = await obtenerToken();

    const response = await fetch(url.startsWith("https://") ? url : GRAPH + url, {
      ...opts,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        ...(opts.body ? { "Content-Type": "application/json" } : {}),
        ...(opts.headers || {})
      }
    });

    // Reintento simple si Graph limita las peticiones.
    if ((response.status === 429 || response.status === 503) && intento < 2) {
      const espera = Number(response.headers.get("Retry-After")) || 2;
      await new Promise((r) => setTimeout(r, espera * 1000));
      return graph(url, opts, intento + 1);
    }

    if (!response.ok) {
      const txt = await response.text();
      let detalle = txt;
      try {
        detalle = JSON.parse(txt)?.error?.message || txt;
      } catch {
        // El cuerpo no es JSON: se usa el texto plano.
      }
      throw new Error(`Graph ${response.status}: ${detalle || response.statusText}`);
    }

    if (response.status === 204) return null;

    return (response.headers.get("content-type") || "").includes("application/json")
      ? response.json()
      : response;
  }

  async function paginar(url) {
    const resultados = [];
    let siguiente = url;

    while (siguiente) {
      const pagina = await graph(siguiente);
      resultados.push(...(pagina.value || []));
      siguiente = pagina["@odata.nextLink"] || null;
    }

    return resultados;
  }

  const resolverSitio = (host, ruta) =>
    graph(`/sites/${host}:${ruta.replace(/\/+$/, "")}?$select=id,displayName,webUrl`);

  // =========================================================
  // SHAREPOINT Y MICROSOFT LISTS
  // =========================================================

  function buscarColumna(nombres, soloEscribibles = false) {
    const buscados = nombres.map(normCol);

    return S.columnas.find(
      (c) =>
        (!soloEscribibles || (!c.hidden && !c.readOnly)) &&
        (buscados.includes(normCol(c.name)) || buscados.includes(normCol(c.displayName)))
    );
  }

  async function encontrarLista() {
    const sitios = [...new Set([S.sitioLista?.id, S.sitioDocs?.id].filter(Boolean))];

    const respuestas = await Promise.all(
      sitios.map((id) =>
        graph(`/sites/${id}/lists?$select=id,displayName`)
          .then((r) => ({ id, listas: r.value || [] }))
          .catch(() => ({ id, listas: [] }))
      )
    );

    for (const { id, listas } of respuestas) {
      const lista = listas.find(
        (l) => normalizar(l.displayName) === normalizar(CFG.lista.name)
      );

      if (lista) {
        S.listaId = lista.id;
        S.sitioLista = { id };
        log("Lista encontrada:", lista.displayName);
        return true;
      }
    }

    warn(`No se encontró la lista "${CFG.lista.name}".`);
    return false;
  }

  async function cargarColumnas() {
    if (!S.listaId || !S.sitioLista) return;

    S.columnas = await paginar(
      `/sites/${S.sitioLista.id}/lists/${S.listaId}/columns?$select=name,displayName,hidden,readOnly`
    );
  }

  async function prepararSharePoint() {
    const [sitioDocs, sitioLista] = await Promise.all([
      resolverSitio(CFG.docs.host, CFG.docs.site),
      resolverSitio(CFG.lista.host, CFG.lista.site).catch((e) => {
        warn("No se resolvió el sitio de la lista:", e.message);
        return null;
      })
    ]);

    S.sitioDocs = sitioDocs;
    S.sitioLista = sitioLista;

    const drives = await paginar(`/sites/${sitioDocs.id}/drives?$select=id,name,driveType`);

    const drive =
      drives.find((d) => /documentos compartidos|shared documents|documentos/i.test(d.name)) ||
      drives.find((d) => d.driveType === "documentLibrary") ||
      drives[0];

    if (!drive) throw new Error("No se encontró una biblioteca documental.");
    S.driveId = drive.id;

    if (await encontrarLista()) {
      await cargarColumnas();
    } else {
      warn("No se podrán guardar conversaciones en Microsoft Lists.");
    }
  }

  async function encontrarCarpeta() {
    const nombre = CFG.docs.folder;

    try {
      const carpeta = await graph(
        `/drives/${S.driveId}/root:/${encodeURIComponent(nombre)}?$select=id,name,folder`
      );
      if (carpeta.folder) return carpeta;
    } catch {
      warn("No se encontró la ruta exacta de la carpeta; se busca en la raíz.");
    }

    const raiz = await paginar(`/drives/${S.driveId}/root/children?$select=id,name,folder`);
    const carpeta = raiz.find((i) => i.folder && normalizar(i.name) === normalizar(nombre));

    if (!carpeta) throw new Error(`No se encontró la carpeta "${nombre}".`);
    return carpeta;
  }

  // Recorre subcarpetas en paralelo.
  async function listarPDF(carpetaId, ruta = "") {
    const items = await paginar(
      `/drives/${S.driveId}/items/${carpetaId}/children?$select=id,name,file,folder,webUrl,eTag`
    );

    const subcarpetas = items.filter((i) => i.folder);
    const archivos = items
      .filter((i) => i.file && /\.pdf$/i.test(i.name))
      .map((i) => ({
        id: i.id,
        nombre: i.name,
        ruta: ruta ? `${ruta}/${i.name}` : i.name,
        webUrl: i.webUrl,
        eTag: i.eTag
      }));

    const anidados = await Promise.all(
      subcarpetas.map((c) => listarPDF(c.id, ruta ? `${ruta}/${c.name}` : c.name))
    );

    return [...archivos, ...anidados.flat()];
  }

  // =========================================================
  // EXTRACCIÓN DE PDF
  // =========================================================

  async function descargarBytes(id) {
    const accessToken = await obtenerToken();

    const response = await fetch(`${GRAPH}/drives/${S.driveId}/items/${id}/content`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!response.ok) throw new Error(`No se pudo descargar el PDF (${response.status}).`);
    return new Uint8Array(await response.arrayBuffer());
  }

  // Divide el texto en fragmentos cortando preferentemente al final de una oración.
  function fragmentar(texto, pagina) {
    const limpio = String(texto || "").replace(/\s+/g, " ").trim();
    if (!limpio) return [];

    const { chunkChars, overlap } = CFG.recuperacion;
    const salida = [];
    let inicio = 0;

    while (inicio < limpio.length) {
      let fin = Math.min(inicio + chunkChars, limpio.length);

      if (fin < limpio.length) {
        const corte = Math.max(limpio.lastIndexOf(". ", fin), limpio.lastIndexOf("; ", fin));
        if (corte > inicio + chunkChars * 0.55) fin = corte + 1;
      }

      const contenido = limpio.slice(inicio, fin).trim();
      if (contenido.length > 30) salida.push({ p: pagina, t: contenido });

      if (fin >= limpio.length) break;
      inicio = Math.max(fin - overlap, inicio + 1);
    }

    return salida;
  }

  async function extraerPDF(archivo) {
    const bytes = await descargarBytes(archivo.id);
    const pdf = await S.pdfJs.getDocument({ data: bytes }).promise;
    const limite = Math.min(pdf.numPages, CFG.maxPaginasPdf);

    const frags = [];

    // Las páginas se leen en lotes pequeños para ganar velocidad sin saturar el navegador.
    for (let inicio = 1; inicio <= limite; inicio += 5) {
      const fin = Math.min(inicio + 4, limite);
      const lote = [];

      for (let n = inicio; n <= fin; n++) {
        lote.push(
          pdf.getPage(n).then((pagina) => pagina.getTextContent().then((c) => ({
            n,
            texto: c.items.map((i) => i.str || "").join(" ")
          })))
        );
      }

      for (const { n, texto } of await Promise.all(lote)) {
        frags.push(...fragmentar(texto, n));
      }
    }

    pdf.destroy?.();

    return { paginas: pdf.numPages, frags };
  }

  // =========================================================
  // ÍNDICE DE BÚSQUEDA (BM25 ligero)
  // =========================================================

  // Tokeniza, quita tildes, stopwords y plural simple.
  function tokens(texto) {
    const salida = [];

    for (const palabra of norm(texto).split(/[^a-z0-9]+/)) {
      if (palabra.length < 2 || STOP.has(palabra) || /^\d+$/.test(palabra)) continue;
      salida.push(palabra.length > 4 && palabra.endsWith("s") ? palabra.slice(0, -1) : palabra);
    }

    return salida;
  }

  function frecuencias(lista) {
    const mapa = new Map();
    for (const t of lista) mapa.set(t, (mapa.get(t) || 0) + 1);
    return mapa;
  }

  // entradas: [{ doc, page, text }]
  function construirIndice(entradas) {
    const frags = entradas.map((e) => {
      const tks = tokens(e.text);
      return {
        doc: e.doc,
        page: e.page,
        text: e.text,
        // Nombre del documento incluido en el conteo para dar peso a títulos.
        tf: frecuencias([...tokens(e.doc), ...tokens(e.doc), ...tks]),
        len: tks.length + tokens(e.doc).length * 2,
        clave: ` ${tks.join(" ")} `
      };
    });

    const df = new Map();
    for (const f of frags) {
      for (const t of f.tf.keys()) df.set(t, (df.get(t) || 0) + 1);
    }

    const avg = frags.reduce((s, f) => s + f.len, 0) / (frags.length || 1);

    return { frags, df, N: frags.length, avg };
  }

  // Devuelve los fragmentos más relevantes, máximo uno por página.
  function buscar(consulta) {
    const idx = S.indice;
    if (!idx?.N) return [];

    const lista = tokens(consulta);
    const unicos = [...new Set(lista)];
    if (!unicos.length) return [];

    const bigramas = [];
    for (let i = 0; i < lista.length - 1; i++) {
      bigramas.push(` ${lista[i]} ${lista[i + 1]} `);
    }

    const K1 = 1.2;
    const B = 0.75;

    const idf = new Map(
      unicos.map((t) => {
        const df = idx.df.get(t) || 0;
        return [t, Math.log(1 + (idx.N - df + 0.5) / (df + 0.5))];
      })
    );

    const puntuados = [];

    for (const f of idx.frags) {
      let puntos = 0;
      let coincidencias = 0;

      for (const t of unicos) {
        const tf = f.tf.get(t);
        if (!tf) continue;

        coincidencias++;
        puntos +=
          (idf.get(t) * (tf * (K1 + 1))) /
          (tf + K1 * (1 - B + (B * f.len) / idx.avg));
      }

      if (!coincidencias) continue;

      // Premia fragmentos que contienen varias palabras de la consulta.
      puntos *= 0.5 + (0.5 * coincidencias) / unicos.length;

      for (const b of bigramas) {
        if (f.clave.includes(b)) puntos += 1.5;
      }

      puntuados.push({ f, puntos });
    }

    puntuados.sort((a, b) => b.puntos - a.puntos);

    const top = puntuados.slice(0, 15);
    if (!top.length || top[0].puntos < 1) return [];

    const umbral = top[0].puntos * 0.4;
    const paginasUsadas = new Set();
    const elegidos = [];

    for (const { f, puntos } of top) {
      if (puntos < umbral) break;

      const clave = `${f.doc}|${f.page}`;
      if (paginasUsadas.has(clave)) continue;

      paginasUsadas.add(clave);
      elegidos.push(f);

      if (elegidos.length >= CFG.recuperacion.topK) break;
    }

    return elegidos;
  }

  function contextoDe(fragmentos) {
    let total = 0;
    const partes = [];
    const fuentes = [];

    fragmentos.forEach((f, i) => {
      const ref = `${f.doc}, pág. ${f.page}`;

      if (total + f.text.length > CFG.recuperacion.maxChars && partes.length) return;

      partes.push(`[Fuente ${i + 1}: ${ref}]\n${f.text}`);
      total += f.text.length;
      fuentes.push(ref);
    });

    return {
      contexto: partes.join("\n\n"),
      fuentes: [...new Set(fuentes)]
    };
  }

  // =========================================================
  // CONTEXTO DE LA CONVERSACIÓN
  // =========================================================

  const ultimoTurno = () => S.historial.at(-1) || null;

  // Si el mensaje es corto o es una continuación ("y si no funciona", "explícame más"),
  // se busca junto con la pregunta anterior para mantener el tema.
  function consultaParaBusqueda(pregunta) {
    const anterior = ultimoTurno();
    if (!anterior) return pregunta;

    const corta = tokens(pregunta).length <= 3;
    const continuacion = /^(y|pero|entonces|eso|esto|si|por que|porque|explica|amplia|detalla|y ahora|y despues|y luego)\b/.test(
      norm(pregunta).trim()
    );

    return corta || continuacion ? `${anterior.pregunta} ${pregunta}` : pregunta;
  }

  // =========================================================
  // IA (OLLAMA) CON STREAMING
  // =========================================================

  function construirPrompt(contexto) {
    return `Eres el asistente interno de soporte de ALFERZA. Atiendes TI, Microsoft 365, Outlook, Teams, equipos, redes, accesos y procedimientos internos.

FORMA DE RESPONDER
- Responde directamente. La primera frase debe ser la respuesta o la acción a seguir.
- Sin relleno: nada de "¡Claro!", "Con gusto", "Espero que te sirva" ni repetir la pregunta.
- Sin emojis. No saludes salvo que el usuario salude.
- Breve por defecto (2 a 6 líneas). Si pide pasos, usa una lista numerada corta y concreta. Si pide más detalle, amplía.
- Responde en el idioma del usuario; por defecto, en español.
- Mantén el hilo: si dice "eso", "y si no funciona" o "¿y después?", aplícalo al tema anterior. Si cambia de tema, sigue el nuevo.
- Si falta un dato clave para ayudar (qué sistema, qué mensaje de error), pídelo en una sola frase al final.

FUENTES
- Si hay CONTEXTO DOCUMENTAL que responde la pregunta, basa los procedimientos internos en él.
- Si no hay contexto o no cubre la pregunta, da orientación técnica general y aclara en una frase que no es un procedimiento interno confirmado.
- Nunca inventes políticas, nombres de personas, extensiones, rutas, enlaces ni acciones realizadas. No digas que revisaste, cambiaste o reiniciaste algo.

TEMAS AJENOS AL TRABAJO
- Puedes responder preguntas generales de forma breve. Si no tienen relación con el trabajo, responde en una o dos frases y ofrece ayudar con TI o procedimientos internos.

CONTEXTO DOCUMENTAL:
${contexto || "(Sin fragmentos relevantes para esta consulta.)"}`;
  }

  async function consultarIA(pregunta, contexto, historial, alParcial) {
    const mensajes = [{ role: "system", content: construirPrompt(contexto) }];

    for (const item of historial.slice(-CFG.historial.turnosIA)) {
      mensajes.push({ role: "user", content: recortar(item.pregunta, CFG.historial.recorteIA) });
      mensajes.push({ role: "assistant", content: recortar(item.respuesta, CFG.historial.recorteIA) });
    }

    mensajes.push({ role: "user", content: recortar(pregunta, 900) });

    const controller = new AbortController();
    let temporizador = setTimeout(() => controller.abort(), CFG.ollama.timeoutMs);
    const reiniciarTemporizador = () => {
      clearTimeout(temporizador);
      temporizador = setTimeout(() => controller.abort(), CFG.ollama.stallMs);
    };

    let texto = "";
    let buffer = "";

    const procesarLinea = (linea) => {
      if (!linea.trim()) return;
      const data = JSON.parse(linea);
      if (data.error) throw new Error(data.error);

      const parte = data.message?.content || "";
      if (parte) {
        texto += parte;
        alParcial?.(texto);
      }
    };

    try {
      const response = await fetch(CFG.ollama.url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: CFG.ollama.model,
          messages: mensajes,
          stream: true,
          keep_alive: CFG.ollama.keepAlive,
          options: {
            temperature: CFG.ollama.temperature,
            num_ctx: CFG.ollama.numCtx,
            num_predict: CFG.ollama.numPredict
          }
        }),
        signal: controller.signal
      });

      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.error || `Ollama respondió con HTTP ${response.status}.`);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder();

      while (true) {
        const { value, done } = await reader.read();
        if (done) break;

        reiniciarTemporizador();
        buffer += decoder.decode(value, { stream: true });

        let salto;
        while ((salto = buffer.indexOf("\n")) >= 0) {
          const linea = buffer.slice(0, salto);
          buffer = buffer.slice(salto + 1);
          procesarLinea(linea);
        }
      }

      procesarLinea(buffer);
    } catch (e) {
      if (e.name === "AbortError") {
        throw new Error("La IA tardó demasiado en responder. Inténtalo de nuevo.");
      }
      if (e instanceof TypeError) {
        throw new Error(
          "No se pudo conectar con Ollama. Revisa que el equipo servidor esté encendido y en la red de ALFERZA."
        );
      }
      throw e;
    } finally {
      clearTimeout(temporizador);
    }

    texto = texto.trim();
    if (!texto) throw new Error("La IA devolvió una respuesta vacía.");
    return texto;
  }

  // Carga el modelo en memoria al iniciar; así la primera consulta del usuario no espera la carga.
  function calentarIA() {
    fetch(CFG.ollama.url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: CFG.ollama.model,
        stream: false,
        keep_alive: CFG.ollama.keepAlive,
        messages: [{ role: "user", content: "ok" }],
        options: { num_predict: 1, num_ctx: CFG.ollama.numCtx }
      })
    }).catch(() => {});
  }

  // =========================================================
  // DOCUMENTOS: CARGA CON CACHÉ INCREMENTAL
  // Cada PDF se identifica por su eTag. Solo se reprocesan los que cambiaron.
  // =========================================================

  async function cargarDocumentos() {
    estadoCarga("Revisando documentos de ALFERZA…");

    const carpeta = await encontrarCarpeta();
    const archivos = await listarPDF(carpeta.id);

    if (!archivos.length) throw new Error(`No hay PDF en "${CFG.docs.folder}".`);

    const cache = await Promise.all(
      archivos.map((a) => idbGet(`pdf:${a.id}`).catch(() => null))
    );

    const docs = [];
    const entradas = [];
    const pendientes = [];

    archivos.forEach((archivo, i) => {
      const guardado = cache[i];

      if (guardado && guardado.eTag === archivo.eTag && Array.isArray(guardado.frags)) {
        docs.push({ nombre: archivo.nombre, paginas: guardado.paginas, fragmentos: guardado.frags.length });
        entradas.push(...guardado.frags.map((f) => ({ doc: archivo.nombre, page: f.p, text: f.t })));
      } else {
        pendientes.push(archivo);
      }
    });

    if (pendientes.length) {
      await cargarPdfJs();

      let listos = 0;
      estadoCarga(`Procesando ${pendientes.length} documento(s) nuevo(s)…`);

      await enParalelo(pendientes, CFG.concurrencia, async (archivo) => {
        try {
          const resultado = await extraerPDF(archivo);

          docs.push({
            nombre: archivo.nombre,
            paginas: resultado.paginas,
            fragmentos: resultado.frags.length
          });
          entradas.push(...resultado.frags.map((f) => ({ doc: archivo.nombre, page: f.p, text: f.t })));

          idbSet(`pdf:${archivo.id}`, {
            eTag: archivo.eTag,
            paginas: resultado.paginas,
            frags: resultado.frags
          }).catch((e) => warn("No se pudo guardar caché:", e.message));
        } catch (e) {
          warn(`No se pudo leer ${archivo.nombre}:`, e.message);
        }

        listos++;
        estadoCarga(`Procesando documentos ${listos}/${pendientes.length}…`);
      });
    }

    if (!entradas.length) throw new Error("No se pudo extraer texto de los PDF.");

    S.docs = docs;
    S.indice = construirIndice(entradas);
    log(`Índice listo: ${S.indice.N} fragmentos de ${docs.length} documento(s).`);
  }

  // =========================================================
  // HISTORIAL
  // =========================================================

  function iniciarHistorialLocal() {
    S.convId = sessionStorage.getItem(LS_KEYS.conv) || uid();
    sessionStorage.setItem(LS_KEYS.conv, S.convId);

    const guardado = leerSesion(LS_KEYS.history, []);

    S.historial = (Array.isArray(guardado) ? guardado : [])
      .filter((i) => typeof i?.pregunta === "string" && typeof i?.respuesta === "string")
      .slice(-CFG.historial.max);
  }

  function guardarHistorialLocal() {
    guardarSesion(LS_KEYS.history, S.historial.slice(-CFG.historial.max));
  }

  function registrarIntercambio(pregunta, respuesta, fuentes = []) {
    const item = { pregunta, respuesta, fuentes, fecha: new Date().toISOString() };

    S.historial = [...S.historial, item].slice(-CFG.historial.max);
    guardarHistorialLocal();

    guardarIntercambio(item).catch(() => {});
    return item;
  }

  // =========================================================
  // MICROSOFT LISTS
  // =========================================================

  async function guardarIntercambio(item) {
    if (!S.listaId || !S.sitioLista) return false;

    try {
      const campos = {};

      const asignar = (nombres, valor) => {
        const columna = buscarColumna(nombres, true);
        if (columna && valor != null && valor !== "") campos[columna.name] = String(valor);
      };

      const titulo = buscarColumna(["Title", "Título"], true);
      if (titulo) campos[titulo.name] = `Chat ALFERZA - ${new Date().toLocaleString("es-PE")}`;

      asignar(["ConversacionId", "ConversationId", "IdConversacion"], S.convId);
      asignar(["Usuario", "Correo", "Email", "UsuarioCorreo"], S.user);
      asignar(["Pregunta", "Consulta", "Mensaje"], item.pregunta);
      asignar(["Respuesta", "RespuestaBot"], item.respuesta);
      asignar(["Fuentes", "Documentos", "Referencias"], (item.fuentes || []).join("; "));
      asignar(["Fecha", "FechaHora", "FechaCreacion"], item.fecha);

      if (!Object.keys(campos).length) return false;

      await graph(`/sites/${S.sitioLista.id}/lists/${S.listaId}/items`, {
        method: "POST",
        body: JSON.stringify({ fields: campos })
      });

      return true;
    } catch (e) {
      warn("No se pudo guardar el intercambio en Microsoft Lists:", e.message);
      return false;
    }
  }

  // Solo se consulta Lists si no hay historial local (por ejemplo, tras recargar con la sesión borrada).
  async function cargarHistorialRemotoSiVacio() {
    if (S.historial.length || !S.listaId || !S.sitioLista) return;

    const cConv = buscarColumna(["ConversacionId", "ConversationId", "IdConversacion"]);
    const cUser = buscarColumna(["Usuario", "Correo", "Email", "UsuarioCorreo"]);
    const cPreg = buscarColumna(["Pregunta", "Consulta", "Mensaje"]);
    const cResp = buscarColumna(["Respuesta", "RespuestaBot"]);
    const cFecha = buscarColumna(["Fecha", "FechaHora", "FechaCreacion"]);

    if (!cPreg || !cResp) return;

    try {
      const pagina = await graph(
        `/sites/${S.sitioLista.id}/lists/${S.listaId}/items?$expand=fields&$top=200`
      );

      const propios = (pagina.value || [])
        .map((item) => item.fields || {})
        .filter(
          (f) =>
            (!cUser || normalizar(f[cUser.name]) === normalizar(S.user)) &&
            (!cConv || String(f[cConv.name] || "") === S.convId)
        )
        .map((f) => ({
          pregunta: String(f[cPreg.name] || ""),
          respuesta: String(f[cResp.name] || ""),
          fuentes: [],
          fecha: cFecha ? f[cFecha.name] || "" : ""
        }))
        .filter((i) => i.pregunta && i.respuesta);

      if (propios.length) {
        S.historial = propios.slice(-CFG.historial.max);
        guardarHistorialLocal();
        renderHistorial();
      }
    } catch (e) {
      warn("No se pudo recuperar el historial remoto:", e.message);
    }
  }

  // =========================================================
  // INICIALIZACIÓN
  // Una sola promesa compartida. Si falla, puede reintentarse.
  // =========================================================

  function iniciarChatbot() {
    if (S.ready) return Promise.resolve();

    if (!S.initPromise) {
      S.ultimoIntentoInit = Date.now();

      S.initPromise = (async () => {
        await iniciarMSAL();
        await prepararSharePoint();
        await cargarDocumentos();

        S.ready = true;
        calentarIA();
        cargarHistorialRemotoSiVacio();
        log("Chatbot listo.");
      })()
        .catch((e) => {
          S.initPromise = null;
          throw e;
        })
        .finally(ocultarCarga);
    }

    return S.initPromise;
  }

  // =========================================================
  // FLUJO DE RESPUESTA
  // =========================================================

  async function obtenerRespuesta(pregunta, alParcial) {
    const social = detectarSocial(pregunta);
    if (social) {
      return { respuesta: respuestaSocial(social), fuentes: [] };
    }

    // Reintenta la conexión documental si falló hace más de un minuto.
    if (!S.ready && Date.now() - S.ultimoIntentoInit > CFG.reintentoDocsMs) {
      estadoCarga("Conectando con los documentos de ALFERZA…");
      try {
        await iniciarChatbot();
      } catch (e) {
        warn("Sin documentos disponibles:", e.message);
      }
    } else if (!S.ready && S.initPromise) {
      estadoCarga("Esperando los documentos de ALFERZA…");
      try {
        await S.initPromise;
      } catch {
        // Se responde sin documentos.
      }
    }

    let contexto = "";
    let fuentes = [];

    if (S.indice?.N) {
      estadoCarga("Buscando en los documentos…");

      const combinada = consultaParaBusqueda(pregunta);
      let hallados = buscar(combinada);
      if (!hallados.length && combinada !== pregunta) hallados = buscar(pregunta);

      ({ contexto, fuentes } = contextoDe(hallados));
    }

    estadoCarga("Preparando la respuesta…");

    let primerTexto = true;
    const alRecibirTexto = (texto) => {
      if (primerTexto) {
        primerTexto = false;
        ocultarCarga();
      }
      alParcial?.(texto);
    };

    const respuesta = await consultarIA(pregunta, contexto, S.historial, alRecibirTexto);

    return { respuesta, fuentes };
  }

  async function procesarPregunta(texto) {
    const pregunta = String(texto ?? "").trim();
    if (!pregunta || S.busy) return;

    const entrada = $("#messageInput");

    S.busy = true;
    setControles();

    crearMensaje("user").burbuja.textContent = pregunta;

    if (entrada) {
      entrada.value = "";
      actualizarContador();
    }

    escribiendo(true);

    const msg = crearMensaje("bot");
    msg.burbuja.textContent = "…";
    const render = renderizadorDiferido(msg);

    try {
      const { respuesta, fuentes } = await obtenerRespuesta(pregunta, (parcial) =>
        render.actualizar(parcial)
      );

      render.cancelar();
      pintarBot(msg, respuesta, fuentes);
      registrarIntercambio(pregunta, respuesta, fuentes);
    } catch (e) {
      render.cancelar();
      console.error("[ALFERZA] Error:", e);
      pintarBot(msg, e.message || "No pude completar la consulta. Inténtalo de nuevo.");
    } finally {
      ocultarCarga();
      escribiendo(false);

      S.busy = false;
      setControles();
      entrada?.focus();
    }
  }

  // =========================================================
  // EVENTOS
  // =========================================================

  function conectarEventos() {
    const form = $("#chatForm");
    const entrada = $("#messageInput");

    form?.addEventListener("submit", (e) => {
      e.preventDefault();
      if (entrada) procesarPregunta(entrada.value);
    });

    entrada?.addEventListener("input", actualizarContador);

    entrada?.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        form?.requestSubmit();
      }
    });

    $("#clearChat")?.addEventListener("click", () => {
      S.historial = [];
      S.convId = uid();
      sessionStorage.setItem(LS_KEYS.conv, S.convId);
      guardarHistorialLocal();
      renderHistorial();
    });

    $$(".suggestion-button").forEach((boton) => {
      boton.addEventListener("click", () => {
        procesarPregunta(boton.dataset.question || boton.textContent.trim());
      });
    });

    actualizarContador();
  }

  // =========================================================
  // ARRANQUE
  // La interfaz aparece de inmediato; los documentos cargan en segundo plano.
  // =========================================================

  function iniciar() {
    conectarEventos();
    iniciarHistorialLocal();
    renderHistorial();

    estadoCarga("Conectando con Microsoft 365…");

    iniciarChatbot().catch((e) => {
      console.error("[ALFERZA] Error de inicialización:", e);
      pintarBot(
        crearMensaje("bot"),
        "No pude conectar con los documentos de ALFERZA. Puedo responder con orientación general; cuando se restablezca la conexión, volverán a estar disponibles."
      );
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", iniciar, { once: true });
  } else {
    iniciar();
  }

  // =========================================================
  // DIAGNÓSTICO (consola del navegador)
  // =========================================================

  window.alferzaChatbotDiagnostico = () => ({
    listo: S.ready,
    usuario: S.user,
    listaDisponible: Boolean(S.listaId),
    ollama: CFG.ollama.url,
    modelo: CFG.ollama.model,
    redirectUri: CFG.redirectUri,
    documentos: S.docs,
    totalFragmentos: S.indice?.N || 0,
    mensajesEnHistorial: S.historial.length
  });

  window.alferzaProbarOllama = async () => {
    const respuesta = await consultarIA(
      "Responde solamente: Ollama funciona correctamente.",
      "",
      [],
      null
    );
    console.log("[ALFERZA] Prueba Ollama:", respuesta);
    return respuesta;
  };
})();