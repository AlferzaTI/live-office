"use strict";

(() => {
  // ===== CONFIGURACIÓN =====
  const CFG = {
    clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
    tenantId: "dbab984f-4bb1-4b60-9dff-da59f54acdf1",
    redirectUri: location.hostname === "localhost"
      ? "http://localhost:5500/blank.html"
      : "https://alferzati.github.io/live-office/blank.html",
    scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"],
    docs: { host: "alferzaholding.sharepoint.com", site: "/sites/TI", folder: "Procedimientos-TI" },
    lista: { host: "alferzaholding-my.sharepoint.com", site: "/personal/soporte1_alferza_pe", name: "Conversaciones-CHATBOT" },
    pdfJs: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
    pdfWorker: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js",
    msalJs: "https://alcdn.msauth.net/browser/2.38.1/js/msal-browser.min.js",
    maxPaginasPdf: 300,
    cacheMs: 15 * 60 * 1000,
    maxCtx: 8,
    maxFragmentos: 2,
    maxChars: 900,
    chunkChars: 1000,
    overlap: 160,
    maxHistorial: 50
  };

  const GRAPH = "https://graph.microsoft.com/v1.0";
  const KEYS = { history: "alferza_chat_history_v2", conv: "alferza_chat_conversation_v2", index: "alferza_pdf_index_v2" };
  const STOPWORDS = new Set(("para como donde cuando cual cuales quiero necesito puedo debo tengo hacer sobre desde hasta este esta estos estas con por una uno unos unas los las del que sin hay ser son fue fueron tambien informacion procedimiento procedimientos documento documentos").split(" "));

  const S = {
    msal: null, account: null, user: "", sitioDocs: null, sitioLista: null,
    driveId: null, listaId: null, columnas: [], docs: [], fragmentos: [],
    historial: [], convId: "", ready: false, initing: false, busy: false, pdfJs: null
  };

  const $ = (s) => document.querySelector(s);
  const $$ = (s) => document.querySelectorAll(s);
  const log = (...a) => console.log("[ALFERZA]", ...a);
  const warn = (...a) => console.warn("[ALFERZA]", ...a);

  const normalizar = (t) => String(t || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/[^a-z0-9@._/-]+/g, " ").trim();
  const normCol = (t) => normalizar(t).replace(/\s/g, "");
  const escapar = (t) => String(t ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  const uid = () => crypto.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`;

  const leer = (k, def) => { try { return JSON.parse(sessionStorage.getItem(k)) ?? def; } catch { return def; } };
  const guardar = (k, v) => { try { sessionStorage.setItem(k, JSON.stringify(v)); } catch (e) { warn("sessionStorage:", e); } };

  // ===== INTERFAZ =====
  function mostrar(texto, tipo = "bot", fuentes = []) {
    const cont = $("#messages");
    if (!cont) return;

    const msg = document.createElement("div");
    msg.className = `message ${tipo}-message`;
    msg.dataset.tipo = tipo;
    msg.innerHTML = `<div class="message-content">${escapar(texto)}</div>`;

    if (fuentes.length && tipo === "bot") {
      msg.insertAdjacentHTML("beforeend",
        `<div class="message-sources" style="font-size:12px;opacity:.8;margin-top:8px">Fuentes: ${escapar([...new Set(fuentes)].join(", "))}</div>`);
    }

    cont.appendChild(msg);
    cont.scrollTop = cont.scrollHeight;
  }

  function estadoCarga(texto) {
    const cont = $("#messages");
    if (!cont) return;
    let aviso = $("#alferzaLoadingStatus");
    if (!aviso) {
      aviso = Object.assign(document.createElement("div"), { id: "alferzaLoadingStatus" });
      aviso.style.cssText = "font-size:13px;opacity:.8;padding:8px 12px";
      cont.appendChild(aviso);
    }
    aviso.textContent = texto;
  }

  const ocultarCarga = () => $("#alferzaLoadingStatus")?.remove();
  const escribiendo = (v) => { const i = $("#typingIndicator"); if (i) i.style.display = v ? "" : "none"; };
  const actualizarContador = () => {
    const e = $("#messageInput"), c = $("#characterCount");
    if (e && c) c.textContent = String(e.value.length);
  };
  const setBoton = () => { const b = $("#sendButton"); if (b) b.disabled = S.busy; };

  function renderHistorial() {
    $("#messages")?.replaceChildren();
    if (!S.historial.length) {
      mostrar("Hola. Soy el asistente documental de ALFERZA. ¿Qué procedimiento o información necesitas consultar?");
      return;
    }
    for (const it of S.historial) {
      mostrar(it.pregunta, "user");
      mostrar(it.respuesta, "bot", it.fuentes || []);
    }
  }

  // ===== MSAL Y GRAPH =====
  async function cargarScript(src, global) {
    if (window[global]) return;
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = src;
      s.onload = res;
      s.onerror = () => rej(new Error(`No se pudo cargar ${src}`));
      document.head.appendChild(s);
    });
    if (!window[global]) throw new Error(`${global} no está disponible.`);
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

    await S.msal.initialize?.();
    const redirect = await S.msal.handleRedirectPromise();
    S.account = redirect?.account || S.msal.getAllAccounts()[0] || null;

    if (!S.account) {
      const login = await S.msal.loginPopup({ scopes: CFG.scopes, prompt: "select_account" });
      S.account = login.account;
    }
    if (!S.account) throw new Error("No se pudo identificar la cuenta Microsoft.");

    S.user = S.account.username || S.account.name || "";
    log("Sesión lista:", S.user);
  }

  async function token() {
    try {
      return (await S.msal.acquireTokenSilent({ scopes: CFG.scopes, account: S.account })).accessToken;
    } catch (e) {
      if (e.name === "InteractionRequiredAuthError" || e.errorCode === "interaction_required") {
        return (await S.msal.acquireTokenPopup({ scopes: CFG.scopes, account: S.account })).accessToken;
      }
      throw e;
    }
  }

  async function graph(url, opts = {}) {
    const t = await token();
    const res = await fetch(url.startsWith("https://") ? url : GRAPH + url, {
      ...opts,
      headers: {
        Authorization: `Bearer ${t}`,
        ...(opts.body ? { "Content-Type": "application/json" } : {}),
        ...(opts.headers || {})
      }
    });

    if (!res.ok) {
      const txt = await res.text();
      let detalle = txt;
      try { detalle = JSON.parse(txt)?.error?.message || txt; } catch {}
      throw new Error(`Graph ${res.status}: ${detalle || res.statusText}`);
    }
    if (res.status === 204) return null;
    return (res.headers.get("content-type") || "").includes("application/json") ? res.json() : res;
  }

  async function paginar(url) {
    const out = [];
    let next = url;
    while (next) {
      const p = await graph(next);
      out.push(...(p.value || []));
      next = p["@odata.nextLink"] || null;
    }
    return out;
  }

  const resolverSitio = (host, ruta) =>
    graph(`/sites/${host}:${ruta.replace(/\/+$/, "")}?$select=id,displayName,webUrl`);

  // ===== SHAREPOINT =====
  async function encontrarLista() {
    const sitios = [S.sitioLista?.id, S.sitioDocs?.id].filter(Boolean);
    const unicos = [...new Set(sitios)];

    for (const sitioId of unicos) {
      try {
        const listas = (await graph(`/sites/${sitioId}/lists?$select=id,displayName`)).value || [];
        const hit = listas.find((l) => normalizar(l.displayName) === normalizar(CFG.lista.name));
        if (hit) {
          S.listaId = hit.id;
          S.sitioLista = { id: sitioId };
          log("Lista encontrada:", hit.displayName);
          return true;
        }
      } catch (e) { warn(`No se consultaron listas en ${sitioId}:`, e); }
    }
    S.listaId = null;
    warn(`No se encontró la lista "${CFG.lista.name}".`);
    return false;
  }

  async function cargarColumnas() {
    if (!S.listaId || !S.sitioLista) return;
    S.columnas = await paginar(
      `/sites/${S.sitioLista.id}/lists/${S.listaId}/columns?$select=name,displayName,required,hidden,readOnly`
    );
  }

  async function prepararSharePoint() {
    S.sitioDocs = await resolverSitio(CFG.docs.host, CFG.docs.site);

    const drives = await paginar(`/sites/${S.sitioDocs.id}/drives?$select=id,name,driveType`);
    const drive = drives.find((d) => /documentos compartidos|shared documents|documentos/i.test(d.name))
      || drives.find((d) => d.driveType === "documentLibrary") || drives[0];
    if (!drive) throw new Error("No se encontró una biblioteca documental.");
    S.driveId = drive.id;

    try { S.sitioLista = await resolverSitio(CFG.lista.host, CFG.lista.site); }
    catch (e) { warn("No se resolvió el sitio de la lista:", e); }

    if (await encontrarLista()) await cargarColumnas();
    else warn("El chatbot buscará en PDF, pero no podrá guardar hasta encontrar la lista.");
  }

  async function listarPDF(carpetaId, ruta = "") {
    const items = await paginar(
      `/drives/${S.driveId}/items/${carpetaId}/children?$select=id,name,file,folder,size,webUrl`
    );
    const out = [];
    for (const it of items) {
      const r = ruta ? `${ruta}/${it.name}` : it.name;
      if (it.folder) out.push(...await listarPDF(it.id, r));
      else if (it.file && /\.pdf$/i.test(it.name))
        out.push({ id: it.id, nombre: it.name, ruta: r, webUrl: it.webUrl });
    }
    return out;
  }

  async function encontrarCarpeta() {
    const nombre = CFG.docs.folder;
    try {
      const r = await graph(`/drives/${S.driveId}/root:/${encodeURIComponent(nombre)}?$select=id,name,folder`);
      if (r.folder) return r;
    } catch (e) { warn("Ruta exacta no encontrada:", e.message); }

    const raiz = await paginar(`/drives/${S.driveId}/root/children?$select=id,name,folder`);
    const hit = raiz.find((i) => i.folder && normalizar(i.name) === normalizar(nombre));
    if (!hit) throw new Error(`No se encontró la carpeta "${nombre}".`);
    return hit;
  }

  // ===== PDF =====
  async function cargarPdfJs() {
    if (!window.pdfjsLib) await cargarScript(CFG.pdfJs, "pdfjsLib");
    S.pdfJs = window.pdfjsLib;
    S.pdfJs.GlobalWorkerOptions.workerSrc = CFG.pdfWorker;
  }

  async function descargarBytes(id) {
    const t = await token();
    const res = await fetch(`${GRAPH}/drives/${S.driveId}/items/${id}/content`, {
      headers: { Authorization: `Bearer ${t}` }
    });
    if (!res.ok) throw new Error(`No se pudo descargar el PDF (${res.status}).`);
    return new Uint8Array(await res.arrayBuffer());
  }

  function fragmentar(texto, documento, pagina) {
    const limpio = String(texto || "").replace(/\s+/g, " ").trim();
    if (!limpio) return [];

    const L = CFG.chunkChars;
    const out = [];
    let inicio = 0;

    while (inicio < limpio.length) {
      let fin = Math.min(inicio + L, limpio.length);
      if (fin < limpio.length) {
        const corte = Math.max(limpio.lastIndexOf(". ", fin), limpio.lastIndexOf("; ", fin));
        if (corte > inicio + L * 0.55) fin = corte + 1;
      }
      const contenido = limpio.slice(inicio, fin).trim();
      if (contenido.length > 30) out.push({ documento, pagina, texto: contenido });
      if (fin >= limpio.length) break;
      inicio = Math.max(fin - CFG.overlap, inicio + 1);
    }
    return out;
  }

  async function extraerPDF(archivo) {
    const pdf = await S.pdfJs.getDocument({ data: await descargarBytes(archivo.id) }).promise;
    const limite = Math.min(pdf.numPages, CFG.maxPaginasPdf);
    const fragmentos = [];

    for (let n = 1; n <= limite; n++) {
      const page = await pdf.getPage(n);
      const { items } = await page.getTextContent();
      const texto = items.map((i) => i.str || "").join(" ");
      fragmentos.push(...fragmentar(texto, archivo.nombre, n));
    }

    if (!fragmentos.length) warn(`"${archivo.nombre}" no produjo texto (¿escaneo sin OCR?).`);
    log(`PDF leído: ${archivo.nombre} (${limite}/${pdf.numPages} págs, ${fragmentos.length} fragmentos)`);

    return {
      meta: { id: archivo.id, nombre: archivo.nombre, webUrl: archivo.webUrl, paginas: pdf.numPages, fragmentos: fragmentos.length },
      fragmentos
    };
  }

  function indiceValido(cache, archivos) {
    return cache && Date.now() - cache.timestamp < CFG.cacheMs
      && cache.archivos?.length === archivos.length
      && archivos.every((a) => cache.archivos.some((c) => c.id === a.id && c.nombre === a.nombre))
      && cache.fragmentos?.length > 0;
  }

  async function cargarDocumentos() {
    await cargarPdfJs();
    estadoCarga("Buscando documentos PDF en SharePoint…");

    const carpeta = await encontrarCarpeta();
    const archivos = await listarPDF(carpeta.id);
    if (!archivos.length) throw new Error(`No hay PDF en "${CFG.docs.folder}".`);

    const cache = leer(KEYS.index, null);
    if (indiceValido(cache, archivos)) {
      S.fragmentos = cache.fragmentos;
      S.docs = cache.archivos;
      log("Índice recuperado de caché.");
      return;
    }

    S.fragmentos = [];
    S.docs = [];

    for (let i = 0; i < archivos.length; i++) {
      estadoCarga(`Leyendo PDF ${i + 1} de ${archivos.length}: ${archivos[i].nombre}`);
      try {
        const { meta, fragmentos } = await extraerPDF(archivos[i]);
        S.docs.push(meta);
        S.fragmentos.push(...fragmentos);
      } catch (e) { warn(`Error al leer "${archivos[i].nombre}":`, e); }
    }

    if (!S.fragmentos.length) throw new Error("No se pudo extraer texto de los PDF.");

    guardar(KEYS.index, { timestamp: Date.now(), archivos: S.docs, fragmentos: S.fragmentos });
    log(`Indexación completa: ${S.docs.length}/${archivos.length} PDF, ${S.fragmentos.length} fragmentos.`);
  }

  // ===== HISTORIAL Y CONTEXTO =====
  function iniciarHistorialLocal() {
    S.convId = sessionStorage.getItem(KEYS.conv) || uid();
    sessionStorage.setItem(KEYS.conv, S.convId);

    const guardado = leer(KEYS.history, []);
    S.historial = (Array.isArray(guardado) ? guardado : [])
      .filter((i) => typeof i?.pregunta === "string" && typeof i?.respuesta === "string")
      .slice(-CFG.maxHistorial);
  }

  const guardarHistorialLocal = () => guardar(KEYS.history, S.historial.slice(-CFG.maxHistorial));

  function esSeguimiento(pregunta) {
    const t = normalizar(pregunta);
    const indicadores = [
      "y que pasa", "y si", "y como", "como hago eso", "cual es el procedimiento",
      "ese procedimiento", "ese documento", "esa informacion", "lo anterior",
      "explicame mas", "detalla mas", "continua", "por que", "y despues",
      "que requisitos", "cuales son los pasos", "donde se hace", "quien lo hace",
      "cuanto demora", "que necesito", "y el siguiente paso", "a que te refieres"
    ];
    return t.split(" ").filter(Boolean).length <= 7 || indicadores.some((i) => t.includes(i));
  }

  function consultaEfectiva(pregunta) {
    const ultima = S.historial.slice(-CFG.maxCtx).at(-1);
    if (!ultima || !esSeguimiento(pregunta)) return pregunta;

    const palabras = new Set(normalizar(pregunta).split(/\s+/).filter((p) => p.length > 3));
    const comparte = normalizar(ultima.pregunta).split(/\s+/).filter((p) => p.length > 3)
      .some((p) => palabras.has(p));

    return comparte ? pregunta : `${ultima.pregunta}. Seguimiento del usuario: ${pregunta}`;
  }

  const extraerPalabras = (t) => normalizar(t).split(/\s+/)
    .filter((p) => p.length >= 3 && !STOPWORDS.has(p) && !/^\d+$/.test(p));

  // ===== BÚSQUEDA =====
  function puntuar(fragmento, consulta, palabras) {
    const texto = normalizar(fragmento.texto);
    const nombre = normalizar(fragmento.documento);
    let pts = 0;

    for (const p of palabras) {
      if (texto.includes(p)) pts += 2;
      if (nombre.includes(p)) pts += 4;
    }

    for (const frase of normalizar(consulta).split(/[?.!,;:]+/).map((s) => s.trim()).filter((s) => s.length > 5)) {
      if (texto.includes(frase)) pts += 8;
    }

    if (/\b(objetivo|alcance|responsable|requisitos|procedimiento|pasos|registro|plazo)\b/i.test(fragmento.texto))
      pts += 0.5;

    return pts;
  }

  function buscar(consulta) {
    const palabras = extraerPalabras(consulta);
    if (!palabras.length || !S.fragmentos.length) return [];

    const candidatos = S.fragmentos
      .map((f) => ({ ...f, puntuacion: puntuar(f, consulta, palabras) }))
      .filter((f) => f.puntuacion > 0)
      .sort((a, b) => b.puntuacion - a.puntuacion);

    const vistos = new Set();
    const elegidos = [];
    for (const c of candidatos) {
      const clave = `${c.documento}|${c.pagina}`;
      if (vistos.has(clave)) continue;
      vistos.add(clave);
      elegidos.push(c);
      if (elegidos.length >= CFG.maxFragmentos) break;
    }
    return elegidos;
  }

  function recortar(texto, maximo) {
    const limpio = String(texto || "").replace(/\s+/g, " ").replace(/^[\s:;,.—-]+/, "").trim();
    if (limpio.length <= maximo) return limpio;

    const corte = limpio.slice(0, maximo);
    const fin = Math.max(corte.lastIndexOf(". "), corte.lastIndexOf("; "), corte.lastIndexOf(": "),
      corte.lastIndexOf("? "), corte.lastIndexOf("! "));
    if (fin > maximo * 0.55) return corte.slice(0, fin + 1).trim();

    const espacio = corte.lastIndexOf(" ");
    return corte.slice(0, espacio > 0 ? espacio : maximo).trim() + "…";
  }

  
function responder(consulta) {
  const pregunta = String(consulta || "").trim();
  const texto = normalizar(pregunta);
  const sinRespuesta = (msg) => ({ respuesta: msg, fuentes: [] });

  if (!pregunta) {
    return sinRespuesta("Escribe tu consulta y te ayudo.");
  }

  // 1. Conversación normal: no buscar en los PDF.
  const saludos = [
    /^hola\b/,
    /^buenas\b/,
    /^buenos dias\b/,
    /^buenas tardes\b/,
    /^buenas noches\b/,
    /^hey\b/,
    /^que tal\b/,
    /^habla\b/,
    /^causa\b/,
    /^causita\b/,
    /^como estas\b/,
    /^como te va\b/,
    /^que haces\b/
  ];

  if (saludos.some((patron) => patron.test(texto))) {
    return {
      respuesta: /como estas|como te va|que tal|que haces/.test(texto)
        ? "¡Todo bien, causita! 😄 Aquí listo para ayudarte. ¿Qué necesitas consultar?"
        : "¡Habla, causita! 😄 Soy el asistente de ALFERZA. ¿Qué necesitas?",
      fuentes: []
    };
  }

  // 2. Agradecimientos.
  if (/^(gracias|muchas gracias|te agradezco|genial gracias)\b/.test(texto)) {
    return sinRespuesta("¡De nada! Aquí estoy para ayudarte.");
  }

  // 3. Despedidas.
  if (/^(chau|chao|adios|hasta luego|nos vemos)\b/.test(texto)) {
    return sinRespuesta("¡Listo! Cuando necesites ayuda, aquí estaré.");
  }

  // 4. Buscar información documental solo para consultas reales.
  const resultados = buscar(pregunta);

  if (!resultados.length) {
    return sinRespuesta(
      "No encontré información suficientemente relacionada con tu consulta en los PDF disponibles. Prueba indicando el nombre del procedimiento, el sistema o el problema específico."
    );
  }

  // Evitar responder con fragmentos de coincidencia débil.
  if (resultados[0].puntuacion < 4) {
    return sinRespuesta(
      "No encontré información suficientemente relacionada con esa pregunta. ¿Puedes especificar qué procedimiento, sistema o problema de TI necesitas consultar?"
    );
  }

  const relevantes = resultados.filter(
    (r) => r.puntuacion >= Math.max(4, resultados[0].puntuacion * 0.5)
  );

  const porFragmento = Math.floor(
    CFG.maxChars / Math.max(1, relevantes.length)
  );

  const partes = relevantes.map((r) => ({
    texto: recortar(r.texto, porFragmento),
    ref: `${r.documento}, pág. ${r.pagina}`
  })).filter((p) => p.texto);

  if (!partes.length) {
    return sinRespuesta(
      "No encontré contenido suficiente para responder con seguridad."
    );
  }

  return {
    respuesta: partes.map((p) => p.texto).join("\n\n"),
    fuentes: [...new Set(partes.map((p) => p.ref))]
  };
}

  // ===== GUARDADO Y HISTORIAL REMOTO =====
  async function guardarIntercambio(item) {
    if (!S.listaId || !S.sitioLista) {
      console.error("[ALFERZA GUARDADO] Lista no encontrada.", { esperada: CFG.lista.name });
      return false;
    }

    try {
      await cargarColumnas();

      const campos = {};
      const columna = (nombres) => {
        const buscados = nombres.map(normCol);
        return S.columnas.find((c) => !c.hidden && !c.readOnly &&
          (buscados.includes(normCol(c.name)) || buscados.includes(normCol(c.displayName))));
      };
      const set = (nombres, valor) => {
        const c = columna(nombres);
        if (c && valor != null) { campos[c.name] = String(valor); return true; }
        return false;
      };

      const titulo = columna(["Title", "Título"]);
      if (titulo) campos[titulo.name] = `Chat ALFERZA - ${new Date().toLocaleString("es-PE")}`;

      set(["ConversacionId", "ConversationId", "IdConversacion"], S.convId);
      set(["Usuario", "Correo", "Email", "UsuarioCorreo"], S.user);
      set(["Pregunta", "Consulta", "Mensaje"], item.pregunta);
      set(["Respuesta", "RespuestaBot"], item.respuesta);
      set(["Fuentes", "Documentos", "Referencias"], (item.fuentes || []).join("; "));
      set(["Fecha", "FechaHora", "FechaCreacion"], item.fecha);

      if (!Object.keys(campos).length)
        throw new Error("No se encontró ninguna columna compatible en la lista.");

      await graph(`/sites/${S.sitioLista.id}/lists/${S.listaId}/items`, {
        method: "POST",
        body: JSON.stringify({ fields: campos })
      });
      log("Registro guardado en la lista.");
      return true;
    } catch (e) {
      console.error("[ALFERZA GUARDADO] Error real al guardar:", e);
      mostrar("Aviso: la respuesta se generó, pero no se pudo confirmar su guardado en Microsoft Lists. Revisa la consola.");
      return false;
    }
  }

  async function cargarHistorialRemoto() {
    if (!S.listaId || !S.sitioLista) return;

    const col = (...n) => {
      const b = n.map(normCol);
      return S.columnas.find((c) => b.includes(normCol(c.name)) || b.includes(normCol(c.displayName)));
    };
    const cConv = col("ConversacionId", "ConversationId", "IdConversacion");
    const cUser = col("Usuario", "Correo", "Email", "UsuarioCorreo");
    const cPreg = col("Pregunta", "Consulta", "Mensaje");
    const cResp = col("Respuesta", "RespuestaBot");
    const cFecha = col("Fecha", "FechaHora", "FechaCreacion");

    if (!cPreg || !cResp) {
      warn("La lista no tiene columnas de pregunta/respuesta; se usa historial local.");
      return;
    }

    try {
      const items = await paginar(`/sites/${S.sitioLista.id}/lists/${S.listaId}/items?$expand=fields&$top=200`);
      const propios = items.map((i) => i.fields || {})
        .filter((f) =>
          (!cUser || normalizar(f[cUser.name]) === normalizar(S.user)) &&
          (!cConv || String(f[cConv.name] || "") === S.convId))
        .map((f) => ({
          pregunta: String(f[cPreg.name] || ""),
          respuesta: String(f[cResp.name] || ""),
          fuentes: [],
          fecha: cFecha ? f[cFecha.name] || "" : ""
        }))
        .filter((i) => i.pregunta && i.respuesta);

      if (propios.length) {
        S.historial = propios.slice(-CFG.maxHistorial);
        guardarHistorialLocal();
        renderHistorial();
        log("Historial remoto recuperado:", propios.length);
      }
    } catch (e) {
      warn("No se recuperó el historial remoto:", e);
    }
  }

  // ===== FLUJO =====
  async function iniciarChatbot() {
    if (S.ready) return;
    if (S.initing) {
      while (S.initing) await new Promise((r) => setTimeout(r, 100));
      return;
    }

    S.initing = true;
    try {
      iniciarHistorialLocal();
      await iniciarMSAL();
      await prepararSharePoint();
      await cargarDocumentos();
      await cargarHistorialRemoto();
      S.ready = true;
      if (!S.historial.length) renderHistorial();
    } finally {
      S.initing = false;
      ocultarCarga();
    }
  }

  async function procesarPregunta(texto) {
    const pregunta = String(texto || "").trim();
    if (!pregunta || S.busy) return;

    S.busy = true;
    setBoton();
    mostrar(pregunta, "user");
    escribiendo(true);

    const entrada = $("#messageInput");
    if (entrada) { entrada.value = ""; actualizarContador(); }

    try {
      if (!S.ready) await iniciarChatbot();

      const { respuesta, fuentes } = responder(consultaEfectiva(pregunta));
      const item = { pregunta, respuesta, fuentes, fecha: new Date().toISOString() };

      S.historial = [...S.historial, item].slice(-CFG.maxHistorial);
      guardarHistorialLocal();
      mostrar(respuesta, "bot", fuentes);

      await guardarIntercambio(item);
    } catch (e) {
      console.error("[ALFERZA] Error al responder:", e);
      mostrar("No pude completar la consulta. Revisa la conexión con Microsoft 365 y que los documentos estén disponibles.");
    } finally {
      escribiendo(false);
      S.busy = false;
      setBoton();
      entrada?.focus();
    }
  }

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
      sessionStorage.setItem(KEYS.conv, S.convId);
      guardarHistorialLocal();
      renderHistorial();
    });

    $$(".suggestion-button").forEach((b) =>
      b.addEventListener("click", () => procesarPregunta(b.dataset.question || b.textContent.trim())));

    actualizarContador();
  }

  async function iniciar() {
    conectarEventos();
    estadoCarga("Conectando con Microsoft 365…");
    try {
      await iniciarChatbot();
    } catch (e) {
      console.error("[ALFERZA] Error de inicialización:", e);
      estadoCarga("No se pudo completar la carga. Revisa la consola para ver el error.");
      mostrar("No pude cargar los documentos de ALFERZA. Verifica tu sesión de Microsoft y los permisos de SharePoint.");
    }
  }

  document.readyState === "loading"
    ? document.addEventListener("DOMContentLoaded", iniciar, { once: true })
    : iniciar();

  window.alferzaChatbotDiagnostico = () => ({
    listo: S.ready,
    usuario: S.user,
    listaDisponible: Boolean(S.listaId),
    pdf: S.docs.map((d) => ({ nombre: d.nombre, paginas: d.paginas, fragmentos: d.fragmentos })),
    totalFragmentos: S.fragmentos.length,
    mensajesEnHistorial: S.historial.length
  });
})();