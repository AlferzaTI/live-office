
"use strict";

(() => {
  // =========================================================
  // ALFERZA CHATBOT — búsqueda documental y contexto conversacional
  // Sin backend propio ni OpenAI.
  // =========================================================

  const CONFIG = {
    clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
    tenantId: "dbab984f-4bb1-4b60-9dff-da59f54acdf1",

    redirectUri:
      window.location.hostname === "localhost"
        ? "http://localhost:5500/blank.html"
        : "https://alferzati.github.io/live-office/blank.html",

    scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"],

    documentos: {
      hostname: "alferzaholding.sharepoint.com",
      rutaSitio: "/sites/TI",
      carpeta: "Procedimientos-TI"
    },

    lista: {
      hostname: "alferzaholding-my.sharepoint.com",
      rutaSitio: "/personal/soporte1_alferza_pe",
      nombre: "ConversacionesCHATBOT"
    },

    pdfJsUrl:
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",

    pdfWorkerUrl:
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js",

    maxMensajesContexto: 8,
    maxFragmentosRespuesta: 2,
    maxCaracteresRespuesta: 900,
    maxCaracteresFragmento: 1000,
    maxPaginasPdf: 300,
    cacheMs: 15 * 60 * 1000
  };

  const GRAPH = "https://graph.microsoft.com/v1.0";
  const STORAGE_HISTORY = "alferza_chat_history_v2";
  const STORAGE_CONVERSATION = "alferza_chat_conversation_v2";
  const STORAGE_INDEX = "alferza_pdf_index_v2";

  const estado = {
    msal: null,
    cuenta: null,
    usuario: "",
    sitioDocumentosId: null,
    sitioListaId: null,
    driveId: null,
    listaId: null,
    columnas: [],
    documentos: [],
    fragmentos: [],
    historial: [],
    conversacionId: "",
    inicializado: false,
    inicializando: false,
    enviando: false,
    pdfJs: null,
    ultimaPregunta: "",
    ultimoTema: "",
    cargaCompleta: false
  };

  const $ = (selector) => document.querySelector(selector);

  function normalizar(texto) {
    return String(texto || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9@._/-]+/g, " ")
      .trim();
  }

  function escaparHTML(texto) {
    return String(texto ?? "").replace(/[&<>"']/g, (caracter) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    })[caracter]);
  }

  function generarId() {
    if (window.crypto && typeof crypto.randomUUID === "function") {
      return crypto.randomUUID();
    }

    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function leerJSON(clave, defecto) {
    try {
      const valor = sessionStorage.getItem(clave);
      return valor ? JSON.parse(valor) : defecto;
    } catch {
      return defecto;
    }
  }

  function guardarJSON(clave, valor) {
    try {
      sessionStorage.setItem(clave, JSON.stringify(valor));
    } catch (error) {
      console.warn("[ALFERZA] No se pudo guardar en sessionStorage:", error);
    }
  }

  function fechaISO() {
    return new Date().toISOString();
  }

  function registrar(...datos) {
    console.log("[ALFERZA CHATBOT]", ...datos);
  }

  function avisar(...datos) {
    console.warn("[ALFERZA CHATBOT]", ...datos);
  }

  // ---------------------------------------------------------
  // INTERFAZ
  // ---------------------------------------------------------

  function mostrarMensaje(texto, tipo = "bot", fuentes = []) {
    const contenedor = $("#messages");
    if (!contenedor) {
      console.warn("No se encontró el elemento #messages.");
      return;
    }

    const mensaje = document.createElement("div");
    mensaje.className = `message ${tipo}-message`;
    mensaje.dataset.tipo = tipo;

    const cuerpo = document.createElement("div");
    cuerpo.className = "message-content";
    cuerpo.textContent = texto;
    mensaje.appendChild(cuerpo);

    if (fuentes.length && tipo === "bot") {
      const pie = document.createElement("div");
      pie.className = "message-sources";
      pie.style.fontSize = "12px";
      pie.style.opacity = "0.8";
      pie.style.marginTop = "8px";
      pie.textContent =
        "Fuentes: " + [...new Set(fuentes)].join(", ");
      mensaje.appendChild(pie);
    }

    contenedor.appendChild(mensaje);
    contenedor.scrollTop = contenedor.scrollHeight;
    return mensaje;
  }

  function mostrarEscribiendo(visible) {
    const indicador = $("#typingIndicator");
    if (indicador) {
      indicador.style.display = visible ? "" : "none";
    }
  }

  function actualizarContador() {
    const entrada = $("#messageInput");
    const contador = $("#characterCount");

    if (entrada && contador) {
      contador.textContent = String(entrada.value.length);
    }
  }

  function limpiarInterfazMensajes() {
    const contenedor = $("#messages");
    if (!contenedor) return;

    contenedor.replaceChildren();
  }

  function mostrarEstadoCarga(texto) {
    const contenedor = $("#messages");
    if (!contenedor) return;

    let aviso = $("#alferzaLoadingStatus");

    if (!aviso) {
      aviso = document.createElement("div");
      aviso.id = "alferzaLoadingStatus";
      aviso.style.fontSize = "13px";
      aviso.style.opacity = "0.8";
      aviso.style.padding = "8px 12px";
      contenedor.appendChild(aviso);
    }

    aviso.textContent = texto;
  }

  function ocultarEstadoCarga() {
    $("#alferzaLoadingStatus")?.remove();
  }

  function renderizarHistorial() {
    limpiarInterfazMensajes();

    if (!estado.historial.length) {
      mostrarMensaje(
        "Hola. Soy el asistente documental de ALFERZA. ¿Qué procedimiento o información necesitas consultar?",
        "bot"
      );
      return;
    }

    for (const item of estado.historial) {
      mostrarMensaje(item.pregunta, "user");
      mostrarMensaje(
        item.respuesta,
        "bot",
        Array.isArray(item.fuentes) ? item.fuentes : []
      );
    }
  }

  function establecerEstadoBoton() {
    const boton = $("#sendButton");
    if (boton) {
      boton.disabled = estado.enviando;
    }
  }

  // ---------------------------------------------------------
  // MSAL Y MICROSOFT GRAPH
  // ---------------------------------------------------------

  async function cargarMSAL() {
    if (window.msal) return;

    await new Promise((resolve, reject) => {
      const existente = document.querySelector(
        'script[src*="msal-browser"]'
      );

      if (existente) {
        existente.addEventListener("load", resolve, { once: true });
        existente.addEventListener(
          "error",
          () => reject(new Error("No se pudo cargar MSAL.")),
          { once: true }
        );

        if (window.msal) resolve();
        return;
      }

      const script = document.createElement("script");
      script.src =
        "https://alcdn.msauth.net/browser/2.38.1/js/msal-browser.min.js";
      script.onload = resolve;
      script.onerror = () =>
        reject(new Error("No se pudo cargar la biblioteca MSAL."));
      document.head.appendChild(script);
    });

    if (!window.msal) {
      throw new Error("MSAL no está disponible.");
    }
  }

  async function inicializarMSAL() {
    await cargarMSAL();

    estado.msal = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority:
          `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: CONFIG.redirectUri,
        navigateToLoginRequestUrl: false
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
      }
    });

    if (typeof estado.msal.initialize === "function") {
      await estado.msal.initialize();
    }

    const resultadoRedirect = await estado.msal.handleRedirectPromise();
    const cuentas = estado.msal.getAllAccounts();

    estado.cuenta =
      resultadoRedirect?.account ||
      cuentas[0] ||
      null;

    if (!estado.cuenta) {
      const login = await estado.msal.loginPopup({
        scopes: CONFIG.scopes,
        prompt: "select_account"
      });

      estado.cuenta = login.account;
    }

    if (!estado.cuenta) {
      throw new Error("No se pudo identificar la cuenta Microsoft.");
    }

    estado.usuario =
      estado.cuenta.username ||
      estado.cuenta.name ||
      "";

    registrar("Sesión Microsoft lista:", estado.usuario);
  }

  async function obtenerToken() {
    if (!estado.msal || !estado.cuenta) {
      throw new Error("La sesión de Microsoft no está inicializada.");
    }

    try {
      const resultado = await estado.msal.acquireTokenSilent({
        scopes: CONFIG.scopes,
        account: estado.cuenta
      });

      return resultado.accessToken;
    } catch (error) {
      if (
        error.name === "InteractionRequiredAuthError" ||
        error.errorCode === "interaction_required"
      ) {
        const resultado = await estado.msal.acquireTokenPopup({
          scopes: CONFIG.scopes,
          account: estado.cuenta
        });

        return resultado.accessToken;
      }

      throw error;
    }
  }

  async function graphFetch(url, opciones = {}) {
    const token = await obtenerToken();
    const urlCompleta = url.startsWith("https://")
      ? url
      : `${GRAPH}${url}`;

    const respuesta = await fetch(urlCompleta, {
      ...opciones,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(opciones.body ? { "Content-Type": "application/json" } : {}),
        ...(opciones.headers || {})
      }
    });

    if (!respuesta.ok) {
      const cuerpo = await respuesta.text();
      let detalle = cuerpo;

      try {
        const json = JSON.parse(cuerpo);
        detalle = json?.error?.message || cuerpo;
      } catch {
        // El cuerpo no era JSON.
      }

      throw new Error(
        `Graph ${respuesta.status}: ${detalle || respuesta.statusText}`
      );
    }

    if (respuesta.status === 204) return null;

    const tipo = respuesta.headers.get("content-type") || "";

    if (tipo.includes("application/json")) {
      return respuesta.json();
    }

    return respuesta;
  }

  async function obtenerPaginas(url) {
    const resultados = [];
    let siguiente = url;

    while (siguiente) {
      const pagina = await graphFetch(siguiente);
      resultados.push(...(pagina.value || []));
      siguiente = pagina["@odata.nextLink"] || null;
    }

    return resultados;
  }

  async function resolverSitio(hostname, ruta) {
    const rutaLimpia = ruta.replace(/\/+$/, "");

    return graphFetch(
      `/sites/${hostname}:${rutaLimpia}?$select=id,displayName,webUrl`
    );
  }

  // ---------------------------------------------------------
  // CONFIGURACIÓN DE SHAREPOINT Y LISTS
  // ---------------------------------------------------------

  function normalizarNombreColumna(texto) {
    return normalizar(texto).replace(/\s/g, "");
  }

  async function encontrarLista() {
    const sitios = [];

    if (estado.sitioListaId) {
      sitios.push({
        id: estado.sitioListaId,
        nombre: "sitio configurado para la lista"
      });
    }

    if (
      estado.sitioDocumentosId &&
      !sitios.some((sitio) => sitio.id === estado.sitioDocumentosId)
    ) {
      sitios.push({
        id: estado.sitioDocumentosId,
        nombre: "sitio de documentos"
      });
    }

    const listasDetectadas = [];

    for (const sitio of sitios) {
      try {
        const datos = await graphFetch(
          `/sites/${sitio.id}/lists?$select=id,displayName,webUrl,list`
        );

        const listas = datos.value || [];

        listasDetectadas.push(
          ...listas.map((lista) => ({
            nombre: lista.displayName,
            sitio: sitio.nombre
          }))
        );

        const encontrada = listas.find(
          (lista) =>
            normalizar(lista.displayName) ===
            normalizar(CONFIG.lista.nombre)
        );

        if (encontrada) {
          estado.listaId = encontrada.id;
          estado.sitioListaId = sitio.id;

          registrar(
            "Lista encontrada:",
            encontrada.displayName,
            "en",
            sitio.nombre
          );

          return true;
        }
      } catch (error) {
        avisar(`No se pudieron consultar listas en ${sitio.nombre}:`, error);
      }
    }

    estado.listaId = null;

    avisar(
      `No se encontró la lista "${CONFIG.lista.nombre}". Listas detectadas:`,
      listasDetectadas
    );

    return false;
  }

  async function prepararSharePoint() {
    const sitioDocs = await resolverSitio(
      CONFIG.documentos.hostname,
      CONFIG.documentos.rutaSitio
    );

    estado.sitioDocumentosId = sitioDocs.id;
    registrar("Sitio documental:", sitioDocs.webUrl);

    const drives = await obtenerPaginas(
      `/sites/${estado.sitioDocumentosId}/drives?$select=id,name,driveType,webUrl`
    );

    const drive =
      drives.find((d) =>
        /documentos compartidos|shared documents|documentos/i.test(d.name)
      ) ||
      drives.find((d) => d.driveType === "documentLibrary") ||
      drives[0];

    if (!drive) {
      throw new Error("No se encontró una biblioteca documental en SharePoint.");
    }

    estado.driveId = drive.id;
    registrar("Biblioteca documental:", drive.name);

    try {
      const sitioLista = await resolverSitio(
        CONFIG.lista.hostname,
        CONFIG.lista.rutaSitio
      );

      estado.sitioListaId = sitioLista.id;
      registrar("Sitio de listas:", sitioLista.webUrl);
    } catch (error) {
      avisar(
        "No se pudo resolver el sitio configurado para la lista.",
        error
      );
    }

    const listaEncontrada = await encontrarLista();

    if (listaEncontrada) {
      try {
        const columnas = await obtenerPaginas(
          `/sites/${estado.sitioListaId}/lists/${estado.listaId}/columns?$select=name,displayName,required,hidden,readOnly`
        );

        estado.columnas = columnas;
        registrar(
          "Columnas de la lista:",
          columnas.map((c) => ({
            nombre: c.name,
            displayName: c.displayName,
            required: c.required
          }))
        );
      } catch (error) {
        avisar("No se pudieron leer las columnas de la lista:", error);
        estado.columnas = [];
      }
    } else {
      avisar(
        "El chatbot podrá buscar en los PDF, pero no podrá guardar en la lista hasta encontrarla."
      );
    }
  }

  // ---------------------------------------------------------
  // CARGA DE PDF.JS Y LECTURA DE DOCUMENTOS
  // ---------------------------------------------------------

  async function cargarPdfJs() {
    if (window.pdfjsLib) {
      estado.pdfJs = window.pdfjsLib;
      estado.pdfJs.GlobalWorkerOptions.workerSrc = CONFIG.pdfWorkerUrl;
      return estado.pdfJs;
    }

    await new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = CONFIG.pdfJsUrl;
      script.onload = resolve;
      script.onerror = () =>
        reject(new Error("No se pudo cargar PDF.js."));
      document.head.appendChild(script);
    });

    if (!window.pdfjsLib) {
      throw new Error("PDF.js se cargó, pero pdfjsLib no está disponible.");
    }

    estado.pdfJs = window.pdfjsLib;
    estado.pdfJs.GlobalWorkerOptions.workerSrc = CONFIG.pdfWorkerUrl;

    return estado.pdfJs;
  }

  async function listarArchivosCarpeta(carpetaId, ruta = "") {
    const elementos = [];
    const pagina = await obtenerPaginas(
      `/drives/${estado.driveId}/items/${carpetaId}/children?$select=id,name,file,folder,size,webUrl`
    );

    for (const item of pagina) {
      const rutaActual = ruta
        ? `${ruta}/${item.name}`
        : item.name;

      if (item.folder) {
        const subelementos = await listarArchivosCarpeta(
          item.id,
          rutaActual
        );

        elementos.push(...subelementos);
      } else if (
        item.file &&
        /\.pdf$/i.test(item.name)
      ) {
        elementos.push({
          id: item.id,
          nombre: item.name,
          ruta: rutaActual,
          webUrl: item.webUrl,
          size: item.size || 0
        });
      }
    }

    return elementos;
  }

  async function encontrarCarpetaDocumental() {
    const carpeta = CONFIG.documentos.carpeta;

    const urlCodificada = encodeURIComponent(carpeta).replace(/%2F/g, "/");

    try {
      const resultado = await graphFetch(
        `/drives/${estado.driveId}/root:/${urlCodificada}?$select=id,name,folder,webUrl`
      );

      if (resultado.folder) return resultado;
    } catch (error) {
      avisar("No se encontró la carpeta por su ruta exacta:", error.message);
    }

    const raiz = await obtenerPaginas(
      `/drives/${estado.driveId}/root/children?$select=id,name,folder,webUrl`
    );

    const encontrada = raiz.find(
      (item) =>
        item.folder &&
        normalizar(item.name) === normalizar(carpeta)
    );

    if (!encontrada) {
      throw new Error(
        `No se encontró la carpeta "${carpeta}" en la biblioteca ${estado.driveId}.`
      );
    }

    return encontrada;
  }

  async function descargarPDF(id) {
    const token = await obtenerToken();

    const respuesta = await fetch(
      `${GRAPH}/drives/${estado.driveId}/items/${id}/content`,
      {
        headers: { Authorization: `Bearer ${token}` }
      }
    );

    if (!respuesta.ok) {
      throw new Error(
        `No se pudo descargar el PDF (${respuesta.status}).`
      );
    }

    return new Uint8Array(await respuesta.arrayBuffer());
  }

  function crearFragmentos(texto, documento, pagina) {
    const limpio = String(texto || "")
      .replace(/\s+/g, " ")
      .trim();

    if (!limpio) return [];

    const longitud = CONFIG.maxCaracteresFragmento;
    const solapamiento = 160;
    const fragmentos = [];

    let inicio = 0;

    while (inicio < limpio.length) {
      let fin = Math.min(inicio + longitud, limpio.length);

      if (fin < limpio.length) {
        const punto = limpio.lastIndexOf(". ", fin);
        const salto = limpio.lastIndexOf("; ", fin);
        const corte = Math.max(punto, salto);

        if (corte > inicio + longitud * 0.55) {
          fin = corte + 1;
        }
      }

      const contenido = limpio.slice(inicio, fin).trim();

      if (contenido.length > 30) {
        fragmentos.push({
          documento,
          pagina,
          texto: contenido
        });
      }

      if (fin >= limpio.length) break;

      inicio = Math.max(fin - solapamiento, inicio + 1);
    }

    return fragmentos;
  }

  async function extraerPDF(archivo) {
    const datos = await descargarPDF(archivo.id);
    const pdf = await estado.pdfJs.getDocument({ data: datos }).promise;
    const fragmentos = [];

    const limite = Math.min(
      pdf.numPages,
      CONFIG.maxPaginasPdf
    );

    for (let numero = 1; numero <= limite; numero++) {
      const pagina = await pdf.getPage(numero);
      const contenido = await pagina.getTextContent();

      const texto = contenido.items
        .map((item) => item.str || "")
        .join(" ")
        .replace(/\s+/g, " ")
        .trim();

      fragmentos.push(
        ...crearFragmentos(texto, archivo.nombre, numero)
      );
    }

    registrar(
      `PDF leído: ${archivo.nombre}`,
      `Páginas: ${limite}/${pdf.numPages}`,
      `Fragmentos: ${fragmentos.length}`
    );

    if (!fragmentos.length) {
      avisar(
        `El PDF "${archivo.nombre}" no produjo texto. Podría ser un escaneo sin OCR.`
      );
    }

    return {
      nombre: archivo.nombre,
      id: archivo.id,
      webUrl: archivo.webUrl,
      paginas: pdf.numPages,
      fragmentos
    };
  }

  async function cargarDocumentos() {
    await cargarPdfJs();

    mostrarEstadoCarga("Buscando documentos PDF en SharePoint…");

    const carpeta = await encontrarCarpetaDocumental();
    const archivos = await listarArchivosCarpeta(carpeta.id);

    if (!archivos.length) {
      throw new Error(
        `No se encontraron archivos PDF dentro de "${CONFIG.documentos.carpeta}".`
      );
    }

    registrar(
      "PDF encontrados en SharePoint:",
      archivos.map((a) => a.ruta)
    );

    const cache = leerJSON(STORAGE_INDEX, null);
    const cacheValida =
      cache &&
      cache.timestamp &&
      Date.now() - cache.timestamp < CONFIG.cacheMs &&
      cache.archivos &&
      cache.archivos.length === archivos.length &&
      archivos.every((a) =>
        cache.archivos.some(
          (c) => c.id === a.id && c.nombre === a.nombre
        )
      ) &&
      Array.isArray(cache.fragmentos) &&
      cache.fragmentos.length > 0;

    if (cacheValida) {
      estado.fragmentos = cache.fragmentos;
      estado.documentos = cache.archivos;
      estado.cargaCompleta = true;

      registrar(
        "Índice de PDF recuperado de caché:",
        estado.documentos.map((d) => d.nombre)
      );

      ocultarEstadoCarga();
      return;
    }

    estado.fragmentos = [];
    estado.documentos = [];

    for (let i = 0; i < archivos.length; i++) {
      const archivo = archivos[i];

      mostrarEstadoCarga(
        `Leyendo PDF ${i + 1} de ${archivos.length}: ${archivo.nombre}`
      );

      try {
        const resultado = await extraerPDF(archivo);

        estado.documentos.push({
          id: resultado.id,
          nombre: resultado.nombre,
          webUrl: resultado.webUrl,
          paginas: resultado.paginas,
          fragmentos: resultado.fragmentos.length
        });

        estado.fragmentos.push(...resultado.fragmentos);
      } catch (error) {
        avisar(`Error al leer "${archivo.nombre}":`, error);
      }
    }

    if (!estado.fragmentos.length) {
      throw new Error(
        "No se pudo extraer texto de los PDF. Revisa permisos, contenido y si los archivos requieren OCR."
      );
    }

    estado.cargaCompleta = true;

    guardarJSON(STORAGE_INDEX, {
      timestamp: Date.now(),
      archivos: estado.documentos,
      fragmentos: estado.fragmentos
    });

    registrar("RESUMEN DE INDEXACIÓN", {
      pdfEncontrados: archivos.length,
      pdfLeidos: estado.documentos.length,
      nombres: estado.documentos.map((d) => d.nombre),
      fragmentosTotales: estado.fragmentos.length
    });

    ocultarEstadoCarga();
  }

  // ---------------------------------------------------------
  // HISTORIAL LOCAL Y CONTEXTO CONVERSACIONAL
  // ---------------------------------------------------------

  function inicializarHistorialLocal() {
    estado.conversacionId =
      sessionStorage.getItem(STORAGE_CONVERSATION) ||
      generarId();

    sessionStorage.setItem(
      STORAGE_CONVERSATION,
      estado.conversacionId
    );

    const guardado = leerJSON(STORAGE_HISTORY, []);

    estado.historial = Array.isArray(guardado)
      ? guardado.filter(
          (item) =>
            item &&
            typeof item.pregunta === "string" &&
            typeof item.respuesta === "string"
        )
      : [];

    estado.historial = estado.historial.slice(-50);
  }

  function guardarHistorialLocal() {
    guardarJSON(STORAGE_HISTORY, estado.historial.slice(-50));
  }

  function obtenerContextoConversacion() {
    return estado.historial
      .slice(-CONFIG.maxMensajesContexto)
      .map((item) => ({
        pregunta: item.pregunta,
        respuesta: item.respuesta
      }));
  }

  function esPreguntaDeSeguimiento(pregunta) {
    const texto = normalizar(pregunta);

    const indicadores = [
      "y que pasa",
      "y si",
      "y como",
      "como hago eso",
      "cual es el procedimiento",
      "ese procedimiento",
      "ese documento",
      "esa informacion",
      "lo anterior",
      "explicame mas",
      "detalla mas",
      "continua",
      "por que",
      "y despues",
      "que requisitos",
      "cuales son los pasos",
      "donde se hace",
      "quien lo hace",
      "cuanto demora",
      "que necesito",
      "y el siguiente paso",
      "a que te refieres"
    ];

    return (
      texto.split(" ").filter(Boolean).length <= 7 ||
      indicadores.some((indicador) => texto.includes(indicador))
    );
  }

  function construirConsultaEfectiva(pregunta) {
    const contexto = obtenerContextoConversacion();
    const ultima = contexto[contexto.length - 1];

    if (!ultima || !esPreguntaDeSeguimiento(pregunta)) {
      return pregunta;
    }

    const preguntaNormalizada = normalizar(pregunta);
    const tema = ultima.pregunta;

    const palabrasPregunta = new Set(
      preguntaNormalizada.split(/\s+/).filter((p) => p.length > 3)
    );

    const preguntaAnterior = normalizar(tema)
      .split(/\s+/)
      .filter((p) => p.length > 3);

    const comparteTema = preguntaAnterior.some((p) =>
      palabrasPregunta.has(p)
    );

    if (comparteTema) return pregunta;

    return `${tema}. Seguimiento del usuario: ${pregunta}`;
  }

  function extraerPalabras(texto) {
    const ignorar = new Set([
      "para", "como", "donde", "cuando", "cual", "cuales",
      "quiero", "necesito", "puedo", "debo", "tengo", "hacer",
      "sobre", "desde", "hasta", "este", "esta", "estos",
      "estas", "con", "por", "una", "uno", "unos", "unas",
      "los", "las", "del", "que", "por", "sin", "hay",
      "ser", "son", "fue", "fueron", "tambien", "informacion",
      "procedimiento", "procedimientos", "documento", "documentos"
    ]);

    return normalizar(texto)
      .split(/\s+/)
      .filter(
        (palabra) =>
          palabra.length >= 3 &&
          !ignorar.has(palabra) &&
          !/^\d+$/.test(palabra)
      );
  }

  // ---------------------------------------------------------
  // BÚSQUEDA RELEVANTE EN TODOS LOS PDF
  // ---------------------------------------------------------

  function puntuarFragmento(fragmento, consulta, palabras) {
    const texto = normalizar(fragmento.texto);
    const nombre = normalizar(fragmento.documento);
    let puntos = 0;

    for (const palabra of palabras) {
      if (texto.includes(palabra)) puntos += 2;
      if (nombre.includes(palabra)) puntos += 4;
    }

    const frases = normalizar(consulta)
      .split(/[?.!,;:]+/)
      .map((s) => s.trim())
      .filter((s) => s.length > 5);

    for (const frase of frases) {
      if (texto.includes(frase)) puntos += 8;
    }

    // Las coincidencias en títulos o encabezados tienden a ser útiles.
    if (
      /\b(objetivo|alcance|responsable|requisitos|procedimiento|pasos|registro|plazo)\b/i
        .test(fragmento.texto)
    ) {
      puntos += 0.5;
    }

    return puntos;
  }

  function buscarFragmentos(consulta) {
    const palabras = extraerPalabras(consulta);

    if (!palabras.length || !estado.fragmentos.length) {
      return [];
    }

    const candidatos = estado.fragmentos
      .map((fragmento) => ({
        ...fragmento,
        puntuacion: puntuarFragmento(
          fragmento,
          consulta,
          palabras
        )
      }))
      .filter((item) => item.puntuacion > 0)
      .sort((a, b) => b.puntuacion - a.puntuacion);

    const seleccionados = [];
    const paginasVistas = new Set();

    for (const item of candidatos) {
      const clave = `${item.documento}|${item.pagina}`;

      // Evita llenar la respuesta con varios fragmentos de la misma página.
      if (paginasVistas.has(clave)) continue;

      seleccionados.push(item);
      paginasVistas.add(clave);

      if (
        seleccionados.length >=
        CONFIG.maxFragmentosRespuesta
      ) {
        break;
      }
    }

    return seleccionados;
  }

  function limpiarFragmento(texto) {
    return String(texto || "")
      .replace(/\s+/g, " ")
      .replace(/^[\s:;,.—-]+/, "")
      .trim();
  }

  function recortarEnFrase(texto, maximo) {
    const limpio = limpiarFragmento(texto);

    if (limpio.length <= maximo) return limpio;

    const corte = limpio.slice(0, maximo);
    const ultimaFrase = Math.max(
      corte.lastIndexOf(". "),
      corte.lastIndexOf("; "),
      corte.lastIndexOf(": "),
      corte.lastIndexOf("? "),
      corte.lastIndexOf("! ")
    );

    if (ultimaFrase > maximo * 0.55) {
      return corte.slice(0, ultimaFrase + 1).trim();
    }

    const ultimaPalabra = corte.lastIndexOf(" ");
    return corte.slice(0, ultimaPalabra > 0 ? ultimaPalabra : maximo).trim() + "…";
  }

  function responderDesdeDocumentos(pregunta, consultaEfectiva) {
    const resultados = buscarFragmentos(consultaEfectiva);

    if (!resultados.length) {
      return {
        respuesta:
          "No encontré una respuesta suficientemente clara en los PDF disponibles. Prueba con el nombre del procedimiento, el área o una palabra específica.",
        fuentes: [],
        confianza: 0
      };
    }

    const mejorPuntuacion = resultados[0].puntuacion;

    if (mejorPuntuacion < 2) {
      return {
        respuesta:
          "No encontré información suficiente en los documentos para responder con seguridad. ¿Puedes indicar el procedimiento o tema exacto?",
        fuentes: [],
        confianza: 0
      };
    }

    const fragmentos = resultados
      .map((item) => ({
        texto: recortarEnFrase(
          item.texto,
          Math.floor(
            CONFIG.maxCaracteresRespuesta /
            resultados.length
          )
        ),
        documento: item.documento,
        pagina: item.pagina
      }))
      .filter((item) => item.texto.length > 0);

    let respuesta = fragmentos
      .map((item) => item.texto)
      .join("\n\n");

    if (respuesta.length > CONFIG.maxCaracteresRespuesta) {
      respuesta = recortarEnFrase(
        respuesta,
        CONFIG.maxCaracteresRespuesta
      );
    }

    return {
      respuesta,
      fuentes: fragmentos.map(
        (item) => `${item.documento}, pág. ${item.pagina}`
      ),
      confianza: mejorPuntuacion
    };
  }

  // ---------------------------------------------------------
  // GUARDADO EN MICROSOFT LISTS
  // ---------------------------------------------------------

  function buscarColumna(...nombres) {
    const buscados = nombres.map(normalizarNombreColumna);

    return estado.columnas.find((columna) => {
      const nombre = normalizarNombreColumna(columna.name);
      const display = normalizarNombreColumna(columna.displayName);

      return buscados.includes(nombre) || buscados.includes(display);
    });
  }

  function establecerCampo(fields, nombres, valor) {
    const columna = buscarColumna(...nombres);

    if (columna && !columna.readOnly && !columna.hidden) {
      fields[columna.name] = valor;
      return true;
    }

    return false;
  }

  async function guardarIntercambio(item) {
    if (!estado.listaId || !estado.sitioListaId) {
      avisar(
        "No se guardó en Microsoft Lists porque la lista configurada no fue encontrada."
      );
      return false;
    }

    const fields = {};

    establecerCampo(
      fields,
      ["Title", "Título"],
      `Chat ${new Date().toLocaleString("es-PE")}`.slice(0, 255)
    );

    establecerCampo(
      fields,
      ["ConversacionId", "ConversationId", "IdConversacion"],
      estado.conversacionId
    );

    establecerCampo(
      fields,
      ["Usuario", "Correo", "Email", "UsuarioCorreo"],
      estado.usuario
    );

    establecerCampo(
      fields,
      ["Pregunta", "Consulta", "Mensaje"],
      item.pregunta
    );

    establecerCampo(
      fields,
      ["Respuesta", "RespuestaBot"],
      item.respuesta
    );

    establecerCampo(
      fields,
      ["Fuentes", "Documentos", "Referencias"],
      (item.fuentes || []).join("; ").slice(0, 2000)
    );

    establecerCampo(
      fields,
      ["Fecha", "FechaHora", "FechaCreacion"],
      item.fecha
    );

    if (!Object.keys(fields).length) {
      avisar(
        "La lista existe, pero no se reconocieron columnas compatibles para guardar el mensaje."
      );
      return false;
    }

    try {
      await graphFetch(
        `/sites/${estado.sitioListaId}/lists/${estado.listaId}/items`,
        {
          method: "POST",
          body: JSON.stringify({ fields })
        }
      );

      registrar("Intercambio guardado en Microsoft Lists.");
      return true;
    } catch (error) {
      avisar("No se pudo guardar el intercambio en Microsoft Lists:", error);
      return false;
    }
  }

  async function cargarHistorialRemoto() {
    if (!estado.listaId || !estado.sitioListaId) return;

    const columnaConversacion = buscarColumna(
      "ConversacionId",
      "ConversationId",
      "IdConversacion"
    );

    const columnaUsuario = buscarColumna(
      "Usuario",
      "Correo",
      "Email",
      "UsuarioCorreo"
    );

    const columnaPregunta = buscarColumna(
      "Pregunta",
      "Consulta",
      "Mensaje"
    );

    const columnaRespuesta = buscarColumna(
      "Respuesta",
      "RespuestaBot"
    );

    const columnaFecha = buscarColumna(
      "Fecha",
      "FechaHora",
      "FechaCreacion"
    );

    if (!columnaPregunta || !columnaRespuesta) {
      avisar(
        "La lista no tiene columnas de pregunta/respuesta reconocibles; se utilizará el historial local."
      );
      return;
    }

    try {
      // La consulta no filtra por columnas personalizadas para evitar
      // errores por nombres internos desconocidos. Se valida localmente.
      const items = await obtenerPaginas(
        `/sites/${estado.sitioListaId}/lists/${estado.listaId}/items?$expand=fields&$top=200`
      );

      const propios = items
        .map((item) => item.fields || {})
        .filter((fields) => {
          const mismoUsuario = !columnaUsuario ||
            normalizar(fields[columnaUsuario.name]) ===
              normalizar(estado.usuario);

          const mismaConversacion = !columnaConversacion ||
            String(fields[columnaConversacion.name] || "") ===
              estado.conversacionId;

          return mismoUsuario && mismaConversacion;
        })
        .map((fields) => ({
          pregunta: String(fields[columnaPregunta.name] || ""),
          respuesta: String(fields[columnaRespuesta.name] || ""),
          fuentes: [],
          fecha: columnaFecha
            ? fields[columnaFecha.name] || ""
            : ""
        }))
        .filter((item) => item.pregunta && item.respuesta);

      if (propios.length) {
        estado.historial = propios.slice(-50);
        guardarHistorialLocal();
        renderizarHistorial();

        registrar(
          "Historial remoto recuperado:",
          propios.length,
          "intercambios"
        );
      }
    } catch (error) {
      avisar(
        "No se pudo recuperar el historial remoto. Se mantiene el historial local.",
        error
      );
    }
  }

  // ---------------------------------------------------------
  // FLUJO PRINCIPAL DEL CHAT
  // ---------------------------------------------------------

  async function procesarPregunta(preguntaOriginal) {
    const pregunta = String(preguntaOriginal || "").trim();

    if (!pregunta) return;

    if (estado.enviando) return;

    estado.enviando = true;
    establecerEstadoBoton();

    mostrarMensaje(pregunta, "user");
    mostrarEscribiendo(true);

    const entrada = $("#messageInput");
    if (entrada) {
      entrada.value = "";
      actualizarContador();
    }

    try {
      if (!estado.inicializado) {
        await iniciarChatbot();
      }

      const consultaEfectiva = construirConsultaEfectiva(pregunta);

      const resultado = responderDesdeDocumentos(
        pregunta,
        consultaEfectiva
      );

      const item = {
        pregunta,
        respuesta: resultado.respuesta,
        fuentes: resultado.fuentes,
        fecha: fechaISO()
      };

      estado.historial.push(item);
      estado.historial = estado.historial.slice(-50);

      guardarHistorialLocal();

      mostrarMensaje(
        resultado.respuesta,
        "bot",
        resultado.fuentes
      );

      estado.ultimaPregunta = pregunta;
      estado.ultimoTema = pregunta;

      await guardarIntercambio(item);
    } catch (error) {
      console.error("[ALFERZA] Error al responder:", error);

      const respuestaError =
        "No pude completar la consulta. Revisa la conexión con Microsoft 365 y que los documentos estén disponibles.";

      mostrarMensaje(respuestaError, "bot");
    } finally {
      mostrarEscribiendo(false);
      estado.enviando = false;
      establecerEstadoBoton();
      entrada?.focus();
    }
  }

  async function iniciarChatbot() {
    if (estado.inicializado) return;
    if (estado.inicializando) {
      while (estado.inicializando) {
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
      return;
    }

    estado.inicializando = true;

    try {
      inicializarHistorialLocal();

      if (sessionStorage.getItem("alferza_login") !== "true") {
        avisar(
          "No se encontró la marca alferza_login. Se comprobará igualmente la sesión de Microsoft."
        );
      }

      await inicializarMSAL();
      await prepararSharePoint();
      await cargarDocumentos();
      await cargarHistorialRemoto();

      estado.inicializado = true;

      registrar("Chatbot inicializado.", {
        usuario: estado.usuario,
        documentos: estado.documentos.map((d) => ({
          nombre: d.nombre,
          paginas: d.paginas,
          fragmentos: d.fragmentos
        })),
        fragmentosTotales: estado.fragmentos.length,
        listaDisponible: Boolean(estado.listaId)
      });

      if (!estado.historial.length) {
        renderizarHistorial();
      }
    } finally {
      estado.inicializando = false;
      ocultarEstadoCarga();
    }
  }

  // ---------------------------------------------------------
  // EVENTOS
  // ---------------------------------------------------------

  function conectarEventos() {
    const formulario = $("#chatForm");
    const entrada = $("#messageInput");
    const botonLimpiar = $("#clearChat");

    if (formulario) {
      formulario.addEventListener("submit", async (evento) => {
        evento.preventDefault();

        if (!entrada) return;

        await procesarPregunta(entrada.value);
      });
    }

    if (entrada) {
      entrada.addEventListener("input", actualizarContador);

      entrada.addEventListener("keydown", (evento) => {
        if (
          evento.key === "Enter" &&
          !evento.shiftKey
        ) {
          evento.preventDefault();
          formulario?.requestSubmit();
        }
      });
    }

    if (botonLimpiar) {
      botonLimpiar.addEventListener("click", () => {
        estado.historial = [];
        estado.conversacionId = generarId();
        estado.ultimaPregunta = "";
        estado.ultimoTema = "";

        sessionStorage.setItem(
          STORAGE_CONVERSATION,
          estado.conversacionId
        );

        guardarHistorialLocal();
        renderizarHistorial();
      });
    }

    document.querySelectorAll(".suggestion-button").forEach((boton) => {
      boton.addEventListener("click", async () => {
        const pregunta =
          boton.dataset.question ||
          boton.textContent.trim();

        await procesarPregunta(pregunta);
      });
    });

    actualizarContador();
  }

  // ---------------------------------------------------------
  // INICIO
  // ---------------------------------------------------------

  async function iniciar() {
    conectarEventos();

    mostrarEstadoCarga("Conectando con Microsoft 365…");

    try {
      await iniciarChatbot();
    } catch (error) {
      console.error("[ALFERZA] Error de inicialización:", error);

      mostrarEstadoCarga(
        "No se pudo completar la carga. Revisa la consola para ver el error."
      );

      mostrarMensaje(
        "No pude cargar los documentos de ALFERZA. Verifica tu sesión de Microsoft y los permisos de SharePoint.",
        "bot"
      );
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", iniciar, {
      once: true
    });
  } else {
    iniciar();
  }

  // Diagnóstico opcional desde la consola del navegador.
  window.alferzaChatbotDiagnostico = () => ({
    inicializado: estado.inicializado,
    usuario: estado.usuario,
    listaDisponible: Boolean(estado.listaId),
    cantidadPDF: estado.documentos.length,
    documentos: estado.documentos.map((d) => ({
      nombre: d.nombre,
      paginas: d.paginas,
      fragmentos: d.fragmentos
    })),
    totalFragmentos: estado.fragmentos.length,
    mensajesEnHistorial: estado.historial.length
  });
})();
