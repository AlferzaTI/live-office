"use strict";

(() => {
  const CFG = {
    clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
    tenantId: "dbab984f-4bb1-4b60-9dff-da59f54acdf1",
    redirectUri: ["localhost", "192.168.2.30"].includes(location.hostname)
      ? `${location.origin}/blank.html`
      : "https://alferzati.github.io/live-office/blank.html",
    scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"],
    docs: { host: "alferzaholding.sharepoint.com", site: "/sites/TI", folder: "Procedimientos-TI" },
    lista: { host: "alferzaholding-my.sharepoint.com", site: "/personal/soporte1_alferza_pe", name: "Conversaciones-CHATBOT" },
    ollama: {
      url: "https://LT-0010.AD.ALFERZA/ollama/api/chat",
      model: "llama3:latest",
      timeoutMs: 45000,
      stallMs: 14000,
      temperature: 0.1,
      numCtx: 2048,
      numPredict: 110,
      keepAlive: "24h"
    },
    pdfJs: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
    pdfWorker: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js",
    msalJs: "https://alcdn.msauth.net/browser/2.38.1/js/msal-browser.min.js",
    maxPaginasPdf: 300,
    concurrencia: 4,
    reintentoDocsMs: 60000,
    esperaDocsMs: 4000,
    rec: { topK: 4, maxChars: 1600, chunkChars: 700, overlap: 80, umbral: 0.28 },
    hist: { max: 50, turnosIA: 6, recorteIA: 280 }
  };

  const GRAPH = "https://graph.microsoft.com/v1.0";
  const LS_HIST = "alferza_chat_history_v3";
  const LS_CONV = "alferza_chat_conversation_v3";
  const STOP = new Set(
    ("a al algo algun alguna ante con como cual cuales cuando de del desde donde e el ella ellas ellos en era es esa ese eso esta estas este esto estos fue ha hacer hay la las le les lo los mas me mi mis mucho muy no o para pero por porque puedo que quien quiero se sin sobre son su sus te tener tengo tu tus un una uno unos y ya yo ayuda favor necesito quisiera dime hola gracias buenas buenos dias tambien si").split(" ")
  );

  const SINONIMOS = {
    correo: ["outlook", "email", "mail", "o365"],
    outlook: ["correo", "email", "mail"],
    email: ["outlook", "correo"],
    wifi: ["red", "internet", "wireless", "conexion"],
    internet: ["wifi", "red", "conexion"],
    red: ["wifi", "internet", "lan", "vpn"],
    impresora: ["impresion", "print", "scanner"],
    contrasena: ["password", "clave", "passwd"],
    clave: ["contrasena", "password"],
    vpn: ["forticlient", "fortinet", "tunel"],
    teams: ["reunion", "videollamada"],
    office: ["microsoft365", "word", "excel"],
    ticket: ["soporte", "incidencia", "mesa"],
    usuario: ["cuenta", "acceso", "login"],
    acceso: ["permiso", "cuenta", "login"],
    pc: ["equipo", "computadora", "laptop"],
    computadora: ["equipo", "pc", "laptop"]
  };

  const TEMAS_TI = new Set(
    ("ti soporte helpdesk mesa ticket incidencia procedimiento alferza outlook correo email mail teams office excel word powerpoint onedrive sharepoint windows laptop pc equipo computadora impresora scanner wifi red internet vpn forticlient fortinet dominio usuario cuenta acceso permiso contrasena password clave login mfa autenticacion bitlocker antivirus licencia software instalar desinstalar actualizacion driver monitor teclado mouse auricular dock cisco activo entra azure intune o365 microsoft365 calendario reunion sala reserva sap qin activo directorio active directory bloqueo pin usb disco backup copia seguridad wifi lan wan dhcp dns proxy firewall").split(" ")
  );

  const FUERA_DE_TEMA = /\b(futbol|futbolista|receta|cocinar|cocina|chiste|chistes|novia|novio|amor|politica|presidente|eleccion|crypto|criptomoneda|bitcoin|apuesta|apuestas|porno|pornografia|sexo|horoscopo|pokemon|minecraft|fortnite|netflix|pelicula|serie|cancion|letra de|tarea de|examen de|historia universal|chisme)\b/;

  const NEGACION_FUERA = "Solo puedo ayudarte con temas de soporte TI y procedimientos de ALFERZA. Reformula tu consulta en ese ámbito.";

  const S = {
    msal: null, account: null, user: "", token: null,
    sitioDocs: null, sitioLista: null, driveId: null, listaId: null,
    columnas: [], indice: null, historial: [], convId: "",
    ready: false, initPromise: null, ultimoIntentoInit: 0, busy: false, pdfJs: null
  };

  const $ = (s) => document.querySelector(s);
  const $$ = (s) => document.querySelectorAll(s);
  const log = (...a) => console.log("[ALFERZA]", ...a);
  const warn = (...a) => console.warn("[ALFERZA]", ...a);
  const pausa = (ms) => new Promise((r) => setTimeout(r, ms));

  const norm = (t) => String(t ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  const normalizar = (t) => norm(t).replace(/[^a-z0-9@._/-]+/g, " ").trim();
  const normCol = (t) => normalizar(t).replace(/\s/g, "");
  const uid = () => crypto.randomUUID?.() || `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const recortar = (t, max) => String(t ?? "").slice(0, max);
  const horaActual = () => new Date().toLocaleTimeString("es-PE", { hour: "2-digit", minute: "2-digit" });
  const el = (tag, clase, texto) => {
    const n = document.createElement(tag);
    if (clase) n.className = clase;
    if (texto != null) n.textContent = texto;
    return n;
  };

  const leerSesion = (k, def) => {
    try { return JSON.parse(sessionStorage.getItem(k)) ?? def; } catch { return def; }
  };
  const guardarSesion = (k, v) => {
    try { sessionStorage.setItem(k, JSON.stringify(v)); } catch (e) { warn("sessionStorage:", e); }
  };

  async function enParalelo(items, n, fn) {
    let i = 0;
    await Promise.all(Array.from({ length: Math.min(n, items.length) }, async () => {
      while (i < items.length) {
        const idx = i++;
        await fn(items[idx], idx);
      }
    }));
  }

  // IndexedDB: los PDF procesados son demasiado grandes para sessionStorage.
  let dbPromise = null;
  const abrirDB = () => (dbPromise ??= new Promise((res, rej) => {
    const req = indexedDB.open("alferza-chatbot", 1);
    req.onupgradeneeded = () => req.result.createObjectStore("kv");
    req.onsuccess = () => res(req.result);
    req.onerror = () => rej(req.error);
  }));

  async function idbGet(clave) {
    const db = await abrirDB();
    return new Promise((res, rej) => {
      const req = db.transaction("kv").objectStore("kv").get(clave);
      req.onsuccess = () => res(req.result ?? null);
      req.onerror = () => rej(req.error);
    });
  }

  async function idbSet(clave, valor) {
    const db = await abrirDB();
    return new Promise((res, rej) => {
      const tx = db.transaction("kv", "readwrite");
      tx.objectStore("kv").put(valor, clave);
      tx.oncomplete = () => res();
      tx.onerror = () => rej(tx.error);
    });
  }

  // Convierte **negrita** y `código` en nodos DOM (sin innerHTML).
  function textoConFormato(elemento, contenido) {
    for (const parte of String(contenido ?? "").split(/(\*\*[^*]+\*\*|`[^`]+`)/g)) {
      if (!parte) continue;
      if (parte.length > 4 && parte.startsWith("**") && parte.endsWith("**")) {
        elemento.appendChild(el("strong", null, parte.slice(2, -2)));
      } else if (parte.length > 2 && parte.startsWith("`") && parte.endsWith("`")) {
        elemento.appendChild(el("code", null, parte.slice(1, -1)));
      } else {
        elemento.appendChild(document.createTextNode(parte));
      }
    }
  }

  function formatearRespuesta(texto) {
    const cont = el("div", "response-formatted");
    let lista = null;
    const cerrarLista = () => { if (lista) { cont.appendChild(lista); lista = null; } };

    for (const original of String(texto || "").replace(/\r\n?/g, "\n").split("\n")) {
      const linea = original.trim();
      if (!linea) { cerrarLista(); continue; }

      const titulo = linea.match(/^#{1,3}\s+(.+)$/);
      const item = linea.match(/^(\d+[.)]|[-*•])\s+(.+)$/);

      if (titulo) {
        cerrarLista();
        const h = el("div", "response-heading");
        textoConFormato(h, titulo[1]);
        cont.appendChild(h);
      } else if (item) {
        const tipo = /^\d/.test(item[1]) ? "ol" : "ul";
        if (!lista || lista.tagName.toLowerCase() !== tipo) {
          cerrarLista();
          lista = el(tipo);
        }
        const li = el("li");
        textoConFormato(li, item[2]);
        lista.appendChild(li);
      } else {
        cerrarLista();
        const p = el("p");
        textoConFormato(p, linea);
        cont.appendChild(p);
      }
    }

    cerrarLista();
    return cont;
  }

  // =========================================================
  // INTERFAZ
  // =========================================================

  const scrollAbajo = () => {
    const cont = $("#messages");
    if (cont) cont.scrollTop = cont.scrollHeight;
  };

  function crearMensaje(tipo) {
    const usuario = tipo === "user";
    const fila = el("div", `message-row ${usuario ? "user" : "bot"}`);
    const avatar = el("div", "message-avatar", usuario ? "TÚ" : "AI");
    const cuerpo = el("div", "message-content");
    const burbuja = el("div", "message-bubble");
    const fuentes = el("div", "message-sources");
    fuentes.hidden = true;

    cuerpo.append(
      el("div", "message-name", usuario ? "Tú" : "Asistente ALFERZA"),
      burbuja,
      fuentes,
      el("div", "message-time", horaActual())
    );
    fila.append(...(usuario ? [cuerpo, avatar] : [avatar, cuerpo]));
    $("#messages")?.appendChild(fila);
    scrollAbajo();

    return { burbuja, fuentes };
  }

  function pintarBot(msg, texto, _fuentes = []) {
    msg.burbuja.replaceChildren(formatearRespuesta(texto));
    msg.fuentes.textContent = "";
    msg.fuentes.hidden = true;
    scrollAbajo();
  }

  // Repinta la burbuja como máximo una vez por frame mientras llega el streaming.
  function renderizadorDiferido(msg) {
    let pendiente = null;
    let frame = 0;
    return {
      actualizar(texto) {
        pendiente = texto;
        if (!frame) {
          frame = requestAnimationFrame(() => {
            frame = 0;
            msg.burbuja.replaceChildren(formatearRespuesta(pendiente));
            scrollAbajo();
          });
        }
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
  // SALUDOS (respuesta instantánea, sin IA)
  // =========================================================

  const RESPUESTAS_SOCIALES = {
    saludo: "Hola. ¿En qué te ayudo?",
    agradecimiento: "Perfecto. ¿Necesitas algo más?",
    despedida: "Hasta luego.",
    estado: "Bien, gracias. ¿En qué te ayudo?"
  };

  function detectarSocial(pregunta) {
    const t = norm(pregunta).replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim();
    if (!t || t.length > 40) return "";

    if (/^(hola+|hey|holi|buenas( tardes| noches| dias)?|buen(os)? dias?|saludos|que tal|hola buenas|hola que tal)( a todos| equipo| amigo)?( que tal| como estas| como va)?$/.test(t)) return "saludo";
    if (/^(muchas )?gracias( mil| totales)?$|^(ok|okay|vale|listo|perfecto|genial|excelente|entendido|claro|dale|ya|bien)( gracias)?$/.test(t)) return "agradecimiento";
    if (/^(chau|chao|adios|hasta luego|hasta pronto|nos vemos|bye|me voy)$/.test(t)) return "despedida";
    if (/^(como estas|como va|todo bien|que tal estas|como te va)$/.test(t)) return "estado";
    return "";
  }

  // =========================================================
  // CARGA DE LIBRERÍAS Y AUTENTICACIÓN
  // =========================================================

  async function cargarScript(src, global) {
    if (window[global]) return;
    await new Promise((res, rej) => {
      const script = el("script");
      script.src = src;
      script.onload = res;
      script.onerror = () => rej(new Error(`No se pudo cargar ${src}`));
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

  async function iniciarMSAL() {
    await cargarScript(CFG.msalJs, "msal");

    S.msal = new msal.PublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: `https://login.microsoftonline.com/${CFG.tenantId}`,
        redirectUri: CFG.redirectUri,
        navigateToLoginRequestUrl: false
      },
      cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
    });

    await S.msal.initialize();
    const redirect = await S.msal.handleRedirectPromise();
    S.account = redirect?.account || S.msal.getAllAccounts()[0] || null;

    if (!S.account) {
      const login = await S.msal.loginPopup({ scopes: CFG.scopes, prompt: "select_account" });
      S.account = login.account;
    }
    if (!S.account) throw new Error("No se pudo identificar la cuenta Microsoft.");

    S.user = S.account.username || S.account.name || "";
    log("Sesión Microsoft lista:", S.user);
  }

  // Token en memoria para no pedirlo a MSAL en cada llamada.
  async function obtenerToken() {
    if (S.token && S.token.expira - Date.now() > 60000) return S.token.valor;

    let resultado;
    try {
      resultado = await S.msal.acquireTokenSilent({ scopes: CFG.scopes, account: S.account });
    } catch (e) {
      if (e.name !== "InteractionRequiredAuthError" && e.errorCode !== "interaction_required") throw e;
      resultado = await S.msal.acquireTokenPopup({ scopes: CFG.scopes, account: S.account });
    }

    S.token = { valor: resultado.accessToken, expira: new Date(resultado.expiresOn).getTime() };
    return S.token.valor;
  }

  // =========================================================
  // MICROSOFT GRAPH
  // =========================================================

  async function graph(url, opts = {}, intento = 0) {
    const token = await obtenerToken();
    const response = await fetch(url.startsWith("https://") ? url : GRAPH + url, {
      ...opts,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(opts.body ? { "Content-Type": "application/json" } : {}),
        ...(opts.headers || {})
      }
    });

    if ((response.status === 429 || response.status === 503) && intento < 2) {
      const espera = Number(response.headers.get("Retry-After")) || 2;
      await pausa(espera * 1000);
      return graph(url, opts, intento + 1);
    }

    if (!response.ok) {
      const txt = await response.text();
      let detalle = txt;
      try { detalle = JSON.parse(txt)?.error?.message || txt; } catch {}
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
      const lista = listas.find((l) => normalizar(l.displayName) === normalizar(CFG.lista.name));
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
      S.columnas = await paginar(
        `/sites/${S.sitioLista.id}/lists/${S.listaId}/columns?$select=name,displayName,hidden,readOnly`
      );
    } else {
      warn("No se podrán guardar conversaciones en Microsoft Lists.");
    }
  }

  async function encontrarCarpeta() {
    const nombre = CFG.docs.folder;
    try {
      const carpeta = await graph(`/drives/${S.driveId}/root:/${encodeURIComponent(nombre)}?$select=id,name,folder`);
      if (carpeta.folder) return carpeta;
    } catch {
      warn("No se encontró la ruta exacta de la carpeta; se busca en la raíz.");
    }

    const raiz = await paginar(`/drives/${S.driveId}/root/children?$select=id,name,folder`);
    const carpeta = raiz.find((i) => i.folder && normalizar(i.name) === normalizar(nombre));
    if (!carpeta) throw new Error(`No se encontró la carpeta "${nombre}".`);
    return carpeta;
  }

  // Recorre subcarpetas en paralelo y devuelve todos los PDF.
  async function listarPDF(carpetaId, ruta = "") {
    const items = await paginar(
      `/drives/${S.driveId}/items/${carpetaId}/children?$select=id,name,file,folder,webUrl,eTag`
    );

    const archivos = items
      .filter((i) => i.file && /\.pdf$/i.test(i.name))
      .map((i) => ({ id: i.id, nombre: i.name, ruta: ruta ? `${ruta}/${i.name}` : i.name, eTag: i.eTag }));

    const anidados = await Promise.all(
      items.filter((i) => i.folder).map((c) => listarPDF(c.id, ruta ? `${ruta}/${c.name}` : c.name))
    );

    return [...archivos, ...anidados.flat()];
  }

  // =========================================================
  // EXTRACCIÓN DE PDF
  // =========================================================

  async function descargarBytes(id) {
    const token = await obtenerToken();
    const response = await fetch(`${GRAPH}/drives/${S.driveId}/items/${id}/content`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!response.ok) throw new Error(`No se pudo descargar el PDF (${response.status}).`);
    return new Uint8Array(await response.arrayBuffer());
  }

  // Divide el texto cortando preferentemente al final de una oración.
  function fragmentar(texto, pagina) {
    const limpio = String(texto || "").replace(/\s+/g, " ").trim();
    const { chunkChars, overlap } = CFG.rec;
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

    // Lotes de 5 páginas, cediendo el hilo principal entre lotes para no congelar la página.
    for (let inicio = 1; inicio <= limite; inicio += 5) {
      const fin = Math.min(inicio + 4, limite);
      const paginas = await Promise.all(
        Array.from({ length: fin - inicio + 1 }, (_, k) => inicio + k).map(async (n) => {
          const page = await pdf.getPage(n);
          const c = await page.getTextContent();
          return { n, texto: c.items.map((i) => i.str || "").join(" ") };
        })
      );
      for (const { n, texto } of paginas) frags.push(...fragmentar(texto, n));
      await pausa(0);
    }

    pdf.destroy?.();
    return { paginas: pdf.numPages, frags };
  }

  // =========================================================
  // ÍNDICE DE BÚSQUEDA (BM25 ligero)
  // =========================================================

  function tokens(texto) {
    return norm(texto)
      .split(/[^a-z0-9]+/)
      .filter((p) => p.length >= 2 && !STOP.has(p) && !/^\d+$/.test(p))
      .map((p) => (p.length > 4 && p.endsWith("s") ? p.slice(0, -1) : p));
  }

  function expandirConsulta(texto) {
    const extra = [];
    for (const t of tokens(texto)) {
      const sin = SINONIMOS[t];
      if (sin) extra.push(...sin);
    }
    return extra.length ? `${texto} ${extra.join(" ")}` : texto;
  }

  function esContinuacion(pregunta) {
    const t = norm(pregunta).trim();
    if (tokens(pregunta).length <= 4) return true;
    return /^(y |pero |entonces |eso |esto |si |ok |y si |por que|porque|explica|amplia|detalla|y ahora|y despues|y luego|lo mismo|igual|otra vez|aun|todavia|no funciona|sigue igual|y despues)\b/.test(t);
  }

  function esTemaTI(pregunta) {
    const t = norm(pregunta);
    if (/\b(outlook|teams|wifi|vpn|excel|office|windows|impresora|sharepoint|onedrive|alferza|ticket|soporte|correo|password|contrasena)\b/.test(t)) {
      return true;
    }
    return tokens(pregunta).some((p) => TEMAS_TI.has(p));
  }

  function esFueraDeTema(pregunta) {
    if (esTemaTI(pregunta)) return false;
    if (S.historial.length && esContinuacion(pregunta)) return false;
    return FUERA_DE_TEMA.test(norm(pregunta));
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
      const tkDoc = tokens(e.doc);
      return {
        doc: e.doc,
        page: e.page,
        text: e.text,
        // El nombre del documento cuenta doble para dar peso a los títulos.
        tf: frecuencias([...tkDoc, ...tkDoc, ...tks]),
        len: tks.length + tkDoc.length * 2,
        clave: ` ${tks.join(" ")} `
      };
    });

    const df = new Map();
    for (const f of frags) for (const t of f.tf.keys()) df.set(t, (df.get(t) || 0) + 1);

    const avg = frags.reduce((s, f) => s + f.len, 0) / (frags.length || 1);
    return { frags, df, N: frags.length, avg };
  }

  // Fragmentos más relevantes, máximo uno por página.
  function buscar(consulta) {
    const idx = S.indice;
    if (!idx?.N) return [];

    const lista = tokens(consulta);
    const unicos = [...new Set(lista)];
    if (!unicos.length) return [];

    const bigramas = lista.slice(0, -1).map((t, i) => ` ${t} ${lista[i + 1]} `);
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
        puntos += (idf.get(t) * tf * (K1 + 1)) / (tf + K1 * (1 - B + (B * f.len) / idx.avg));
      }

      if (!coincidencias) continue;
      puntos *= 0.5 + (0.5 * coincidencias) / unicos.length;
      for (const b of bigramas) if (f.clave.includes(b)) puntos += 1.5;

      puntuados.push({ f, puntos });
    }

    puntuados.sort((a, b) => b.puntos - a.puntos);
    const top = puntuados.slice(0, 24);
    if (!top.length || top[0].puntos < 0.7) return [];

    const umbral = top[0].puntos * CFG.rec.umbral;
    const usados = new Set();
    const elegidos = [];

    for (const { f, puntos } of top) {
      if (puntos < umbral) break;
      const clave = `${f.doc}|${f.page}|${f.text.slice(0, 40)}`;
      if (usados.has(clave)) continue;
      usados.add(clave);
      elegidos.push(f);
      if (elegidos.length >= CFG.rec.topK) break;
    }

    return elegidos;
  }

  function contextoDe(fragmentos) {
    let total = 0;
    const partes = [];

    for (const f of fragmentos) {
      if (total + f.text.length > CFG.rec.maxChars && partes.length) break;
      partes.push(f.text);
      total += f.text.length;
    }

    return { contexto: partes.join("\n\n"), fuentes: [] };
  }

  function consultaParaBusqueda(pregunta) {
    const reciente = S.historial.slice(-2);
    if (!reciente.length || !esContinuacion(pregunta)) return expandirConsulta(pregunta);

    const hilo = reciente
      .map((item) => `${item.pregunta} ${recortar(item.respuesta, 120)}`)
      .join(" ");
    return expandirConsulta(`${hilo} ${pregunta}`);
  }

  function limpiarRespuesta(texto) {
    let t = String(texto || "").trim();
    t = t.replace(/^\s*(claro|por supuesto|con gusto|entendido|okey|okay)[,.]?\s+/i, "");
    t = t.replace(/\b(según|segun|en|del|de)\s+(el\s+)?(documento|pdf|archivo|procedimiento|manual)[^\n.]{0,80}/gi, "");
    t = t.replace(/\b(fuente\s*\d+|p[áa]g(?:ina|\.)?\s*\d+|sharepoint|procedimientos-ti)\b/gi, "");
    t = t.replace(/\[fuente[^\]]*\]/gi, "");
    t = t.replace(/\n{3,}/g, "\n\n").trim();

    const lineas = t.split(/\n+/).filter(Boolean);
    if (lineas.length > 6) t = lineas.slice(0, 6).join("\n");
    if (t.length > 650) {
      const corte = t.lastIndexOf(".", 650);
      t = (corte > 200 ? t.slice(0, corte + 1) : t.slice(0, 650)).trim();
    }
    return t;
  }

  // =========================================================
  // IA (OLLAMA) CON STREAMING
  // =========================================================

  function construirPrompt(contexto) {
    return `Eres el asistente de soporte TI interno de ALFERZA.
Responde en español, directo, en 2 a 5 líneas (máximo 90 palabras). Sin saludos, emojis ni relleno.
La primera frase debe resolver la duda o indicar el siguiente paso.
Usa el conocimiento interno para contestar AQUÍ MISMO. No digas “revisa el documento”, no cites PDF, carpetas, páginas ni rutas.
Si el conocimiento cubre la pregunta, síguelo. Si es de TI pero no aparece ahí, da una orientación breve y honesta, sin inventar nombres de sistemas ni acciones ya hechas.
Si la pregunta no tiene nada que ver con TI, Microsoft 365, equipos, redes, accesos o procedimientos de ALFERZA, responde exactamente:
"${NEGACION_FUERA}"
Si el usuario continúa un tema anterior (“eso”, “y si no funciona”, “y luego”), sigue ese hilo.

CONOCIMIENTO INTERNO:
${contexto || "(sin procedimiento específico)"}`;
  }

  async function consultarIA(pregunta, contexto, historial, alParcial) {
    const mensajes = [{ role: "system", content: construirPrompt(contexto) }];
    for (const item of historial.slice(-CFG.hist.turnosIA)) {
      mensajes.push({ role: "user", content: recortar(item.pregunta, CFG.hist.recorteIA) });
      mensajes.push({ role: "assistant", content: recortar(item.respuesta, CFG.hist.recorteIA) });
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
            num_predict: CFG.ollama.numPredict,
            top_p: 0.8,
            repeat_penalty: 1.15,
            stop: ["Usuario:", "User:", "###"]
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
      if (e.name === "AbortError") throw new Error("La IA tardó demasiado en responder. Inténtalo de nuevo.");
      if (e instanceof TypeError) {
        throw new Error("No se pudo conectar con Ollama. Revisa que el equipo servidor esté encendido y en la red de ALFERZA.");
      }
      throw e;
    } finally {
      clearTimeout(temporizador);
    }

    texto = limpiarRespuesta(texto);
    if (!texto) throw new Error("La IA devolvió una respuesta vacía.");
    return texto;
  }

  // Carga el modelo en memoria apenas se abre la página, sin esperar a SharePoint.
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
  // DOCUMENTOS CON CACHÉ INCREMENTAL (por eTag)
  // =========================================================

  async function cargarDocumentos() {
    estadoCarga("Revisando documentos de ALFERZA…");

    const carpeta = await encontrarCarpeta();
    const archivos = await listarPDF(carpeta.id);
    if (!archivos.length) throw new Error(`No hay PDF en "${CFG.docs.folder}".`);

    const cache = await Promise.all(archivos.map((a) => idbGet(`pdf:${a.id}`).catch(() => null)));

    const entradas = [];
    const pendientes = [];

    archivos.forEach((archivo, i) => {
      const guardado = cache[i];
      if (guardado?.eTag === archivo.eTag && Array.isArray(guardado.frags)) {
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
          entradas.push(...resultado.frags.map((f) => ({ doc: archivo.nombre, page: f.p, text: f.t })));
          idbSet(`pdf:${archivo.id}`, { eTag: archivo.eTag, paginas: resultado.paginas, frags: resultado.frags })
            .catch((e) => warn("No se pudo guardar caché:", e.message));
        } catch (e) {
          warn(`No se pudo leer ${archivo.nombre}:`, e.message);
        }
        estadoCarga(`Procesando documentos ${++listos}/${pendientes.length}…`);
      });
    }

    if (!entradas.length) throw new Error("No se pudo extraer texto de los PDF.");

    S.indice = construirIndice(entradas);
    log(`Índice listo: ${S.indice.N} fragmentos.`);
  }

  // =========================================================
  // HISTORIAL
  // =========================================================

  function iniciarHistorialLocal() {
    S.convId = sessionStorage.getItem(LS_CONV) || uid();
    sessionStorage.setItem(LS_CONV, S.convId);

    const guardado = leerSesion(LS_HIST, []);
    S.historial = (Array.isArray(guardado) ? guardado : [])
      .filter((i) => typeof i?.pregunta === "string" && typeof i?.respuesta === "string")
      .slice(-CFG.hist.max);
  }

  const guardarHistorialLocal = () => guardarSesion(LS_HIST, S.historial.slice(-CFG.hist.max));

  function registrarIntercambio(pregunta, respuesta, fuentes = []) {
    const item = { pregunta, respuesta, fuentes, fecha: new Date().toISOString() };
    S.historial = [...S.historial, item].slice(-CFG.hist.max);
    guardarHistorialLocal();
    guardarIntercambio(item).catch(() => {});
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

  // Solo se consulta Lists si no hay historial local (p. ej., tras borrar la sesión).
  async function cargarHistorialRemotoSiVacio() {
    if (S.historial.length || !S.listaId || !S.sitioLista) return;

    const cConv = buscarColumna(["ConversacionId", "ConversationId", "IdConversacion"]);
    const cUser = buscarColumna(["Usuario", "Correo", "Email", "UsuarioCorreo"]);
    const cPreg = buscarColumna(["Pregunta", "Consulta", "Mensaje"]);
    const cResp = buscarColumna(["Respuesta", "RespuestaBot"]);
    const cFecha = buscarColumna(["Fecha", "FechaHora", "FechaCreacion"]);
    if (!cPreg || !cResp) return;

    try {
      const pagina = await graph(`/sites/${S.sitioLista.id}/lists/${S.listaId}/items?$expand=fields&$top=200`);
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
        S.historial = propios.slice(-CFG.hist.max);
        guardarHistorialLocal();
        renderHistorial();
      }
    } catch (e) {
      warn("No se pudo recuperar el historial remoto:", e.message);
    }
  }

  // =========================================================
  // INICIALIZACIÓN Y FLUJO DE RESPUESTA
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

  async function obtenerRespuesta(pregunta, alParcial) {
    const social = detectarSocial(pregunta);
    if (social) return { respuesta: RESPUESTAS_SOCIALES[social], fuentes: [] };

    if (esFueraDeTema(pregunta)) {
      return { respuesta: NEGACION_FUERA, fuentes: [] };
    }

    if (!S.ready && !S.initPromise && Date.now() - S.ultimoIntentoInit > CFG.reintentoDocsMs) {
      iniciarChatbot().catch((e) => warn("Sin documentos disponibles:", e.message));
    }

    if (!S.ready && S.initPromise) {
      await Promise.race([S.initPromise.catch(() => {}), pausa(CFG.esperaDocsMs)]);
    }

    let contexto = "";
    const tBusqueda = performance.now();

    if (S.indice?.N) {
      const combinada = consultaParaBusqueda(pregunta);
      let hallados = buscar(combinada);
      if (!hallados.length) hallados = buscar(expandirConsulta(pregunta));
      if (!hallados.length && combinada !== pregunta) hallados = buscar(pregunta);
      ({ contexto } = contextoDe(hallados));
    }
    log(`Búsqueda: ${(performance.now() - tBusqueda).toFixed(0)} ms`);

    if (!contexto && !esTemaTI(pregunta) && !(S.historial.length && esContinuacion(pregunta))) {
      return { respuesta: NEGACION_FUERA, fuentes: [] };
    }

    estadoCarga("Preparando la respuesta…");

    const tIA = performance.now();
    let primerTexto = true;
    const respuesta = await consultarIA(pregunta, contexto, S.historial, (texto) => {
      if (primerTexto) {
        primerTexto = false;
        log(`Primer token: ${(performance.now() - tIA).toFixed(0)} ms`);
        ocultarCarga();
      }
      alParcial?.(limpiarRespuesta(texto));
    });
    log(`IA total: ${(performance.now() - tIA).toFixed(0)} ms`);

    return { respuesta, fuentes: [] };
  }

  async function procesarPregunta(texto) {
    const pregunta = String(texto ?? "").trim();
    if (!pregunta || S.busy) return;

    const entrada = $("#messageInput");
    S.busy = true;
    reproducirCuphead("espera");
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
      const { respuesta, fuentes } = await obtenerRespuesta(pregunta, (parcial) => render.actualizar(parcial));
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
      reproducirCuphead("normal");
      S.busy = false;
      setControles();
      entrada?.focus();
    }
  }

  // =========================================================
  // EVENTOS Y ARRANQUE
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
      sessionStorage.setItem(LS_CONV, S.convId);
      guardarHistorialLocal();
      renderHistorial();
    });

    $$(".suggestion-button").forEach((boton) => {
      boton.addEventListener("click", () => procesarPregunta(boton.dataset.question || boton.textContent.trim()));
    });

    actualizarContador();
  }

  function iniciar() {
    conectarEventos();
    iniciarHistorialLocal();
    renderHistorial();

    // Primero el modelo: así la primera pregunta no paga la carga en memoria.
    calentarIA();

    estadoCarga("Conectando con Microsoft 365…");

    iniciarChatbot().catch((e) => {
      console.error("[ALFERZA] Error de inicialización:", e);
      pintarBot(
        crearMensaje("bot"),
        "No pude cargar los procedimientos de ALFERZA. Si tu consulta es de TI, inténtalo de nuevo en un momento."
      );
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", iniciar, { once: true });
  } else {
    iniciar();
  }

  window.alferzaChatbotDiagnostico = () => ({
    listo: S.ready,
    usuario: S.user,
    listaDisponible: Boolean(S.listaId),
    ollama: CFG.ollama.url,
    modelo: CFG.ollama.model,
    redirectUri: CFG.redirectUri,
    totalFragmentos: S.indice?.N || 0,
    mensajesEnHistorial: S.historial.length
  });

  window.alferzaProbarOllama = async () => {
    const respuesta = await consultarIA("Responde solamente: Ollama funciona correctamente.", "", [], null);
    console.log("[ALFERZA] Prueba Ollama:", respuesta);
    return respuesta;
  };
})();



/* =========================================================
   ALFERZA LIVE OFFICE - CUPHEAD ANIMADO
   Normal: frames 0001–0022
   Espera: frames 0025–0071
========================================================= */

(() => {
    "use strict";

    const cupheadFrame = document.getElementById("cupheadFrame");

    if (!cupheadFrame) {
        console.warn(
            "[CUPHEAD] No se encontró #cupheadFrame. Revisa el HTML."
        );
        return;
    }

    const CUPHEAD_BASE = "img/Cuphead/Cuphead/";

    const CUPHEAD_ANIMACIONES = {
        normal: {
            inicio: 1,
            fin: 13,
            velocidad: 110
        },
        espera: {
            inicio: 25,
            fin: 71,
            velocidad: 90
        }
    };

    let cupheadTimer = null;
    let cupheadModoActual = null;
    let cupheadNumeroFrame = 1;

    // Precargar los fotogramas para reducir parpadeos.
    const cupheadImagenes = {};

    for (let i = 1; i <= 71; i++) {
        const numero = String(i).padStart(4, "0");
        const imagen = new Image();

        imagen.src =
            `${CUPHEAD_BASE}mc_reg_intro_cuphead_${numero}.png`;

        cupheadImagenes[i] = imagen;
    }

    function reproducirCuphead(modo) {
        const animacion = CUPHEAD_ANIMACIONES[modo];

        if (!animacion) {
            console.warn(`[CUPHEAD] Animación desconocida: ${modo}`);
            return;
        }

        // Detener cualquier animación anterior.
        clearTimeout(cupheadTimer);

        cupheadModoActual = modo;
        cupheadNumeroFrame = animacion.inicio;

        function siguienteFrame() {
            // Evitar que continúe una animación reemplazada.
            if (cupheadModoActual !== modo) return;

            const imagen = cupheadImagenes[cupheadNumeroFrame];

            if (imagen && imagen.complete && imagen.naturalWidth > 0) {
                cupheadFrame.src = imagen.src;
            }

            cupheadNumeroFrame++;

            if (cupheadNumeroFrame > animacion.fin) {
                cupheadNumeroFrame = animacion.inicio;
            }

            cupheadTimer = setTimeout(
                siguienteFrame,
                animacion.velocidad
            );
        }

        siguienteFrame();
    }

    // Permitir que procesarPregunta() controle la animación.
    window.reproducirCuphead = reproducirCuphead;

    // Empezar en modo normal al abrir el chatbot.
    reproducirCuphead("normal");

})();