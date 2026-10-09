/* =========================================================
   CONTROL DE ACCESO AL INVENTARIO TI
========================================================= */

(function controlarAccesoInventario() {
    const inventario = document.getElementById("inventarioTISection");

    if (!inventario) return;

    // Ocultar únicamente el inventario mientras se validan permisos.
    inventario.style.display = "none";

    function actualizarVisibilidadInventario() {
        const permisos = window.alferzaPermisos || {};
        const grupo = String(permisos.grupo || "")
            .trim()
            .toUpperCase();

        // Solo TI puede ver el inventario.
        inventario.style.display = grupo === "TI" ? "" : "none";
    }

    window.addEventListener(
        "alferza:access-granted",
        actualizarVisibilidadInventario
    );

    if (window.alferzaAccessReady &&
        typeof window.alferzaAccessReady.then === "function") {
        window.alferzaAccessReady
            .then(actualizarVisibilidadInventario)
            .catch(() => {
                inventario.style.display = "none";
            });
    } else {
        actualizarVisibilidadInventario();
    }
})();

const msalConfigInfraestructura={auth:{clientId:"5d98417c-74a7-4fab-8f2c-41ac127be696",authority:"https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",redirectUri:"https://AlferzaTI.github.io/live-office/blank.html"},cache:{cacheLocation:"sessionStorage",storeAuthStateInCookie:false}};
const msalInstanceInfraestructura=(typeof msal!=="undefined")?new msal.PublicClientApplication(msalConfigInfraestructura):null;

const SHAREPOINT_HOSTNAME="alferzaholding-my.sharepoint.com";
const SHAREPOINT_SITE="/personal/soporte1_alferza_pe";
const SHAREPOINT_LIST_SERVICIOS="MonitoreoTI";
const SHAREPOINT_LIST_TRABAJADORES="TrabajadoresEquipos";
const RUTA_EQUIPOS="Procedimientos T.I/equipos.json";

const EQUIPOS_POR_PAGINA=40;
const INTERVALO_ACTUALIZACION=120000;

let cuentaActual=null;
let tokenActual=null;
let sitioSharePoint=null;
let driveSharePoint=null;

let listaTrabajadores=null;
let columnasTrabajadores=null;
let trabajadoresData=[];
let trabajadoresPorEquipo={};
let trabajadoresCargados=false;
let cargandoTrabajadores=false;

let equiposData=[];
let equiposFiltrados=[];
let paginaActual=1;

let equipoSeleccionado=null;

const loadingOverlay=document.getElementById("loadingOverlay");
const loadingMessage=document.getElementById("loadingMessage");
const serviciosGrid=document.getElementById("serviciosGrid");
const estadoGeneral=document.getElementById("estadoGeneral");
const equiposTableBody=document.getElementById("equiposTableBody");
const equiposEmpty=document.getElementById("equiposEmpty");
const equiposResultados=document.getElementById("equiposResultados");
const totalEquipos=document.getElementById("totalEquipos");
const equiposConectados=document.getElementById("equiposConectados");
const equiposDesconectados=document.getElementById("equiposDesconectados");
const filtroEstadoEquipo=document.getElementById("filtroEstadoEquipo");
const filtroAreaEquipo=document.getElementById("filtroAreaEquipo");
const buscarEquipo=document.getElementById("buscarEquipo");
const equiposEstado=document.getElementById("equiposEstado");
const lastUpdate=document.getElementById("lastUpdate");

const gestionTrabajadorModal=document.getElementById("gestionTrabajadorModal");
const gestionIP=document.getElementById("gestionIP");
const buscarEquipoIP=document.getElementById("buscarEquipoIP");
const equipoEncontrado=document.getElementById("equipoEncontrado");
const gestionEquipoNombre=document.getElementById("gestionEquipoNombre");
const gestionEquipoDetalle=document.getElementById("gestionEquipoDetalle");
const gestionTrabajador=document.getElementById("gestionTrabajador");
const gestionTrabajadorMensaje=document.getElementById("gestionTrabajadorMensaje");
const btnGestionarTrabajadores=document.getElementById("btnGestionarTrabajadores");
const btnRegistrarTrabajador=document.getElementById("btnRegistrarTrabajador");
const btnEditarTrabajador=document.getElementById("btnEditarTrabajador");
const btnEliminarTrabajador=document.getElementById("btnEliminarTrabajador");
const cerrarGestionTrabajadorModal=document.getElementById("cerrarGestionTrabajadorModal");
const cancelarGestionTrabajador=document.getElementById("cancelarGestionTrabajador");

function escaparHTML(valor){
if(valor===null||valor===undefined){return "";}
return String(valor).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#039;");
}

function mostrarLoading(){}

function ocultarLoading(){
if(loadingOverlay){loadingOverlay.style.display="none";}
}

function obtenerFechaActual(){
return new Date().toLocaleString("es-PE",{dateStyle:"short",timeStyle:"medium"});
}

function normalizarTexto(valor){
return String(valor||"").trim().toLowerCase();
}

function sinTildes(valor){
return String(valor===null||valor===undefined?"":valor)
.replace(/\_x([0-9a-f]{4})\_/gi,(m,h)=>String.fromCharCode(parseInt(h,16)))
.normalize("NFD")
.replace(/[\u0300-\u036f]/g,"")
.toLowerCase()
.trim();
}

function ayudaError(texto){
const t=String(texto);
const ayudas=[
[/AADSTS65001|consent_required/i,"Falta consentimiento de administrador. En Azure AD → Registros de aplicaciones → ALFERZA Live Office → Permisos de API, agrega Sites.Read.All y Sites.ReadWrite.All y concede consentimiento."],
[/AADSTS70011|invalid_scope/i,"Un permiso no existe o no está configurado en la aplicación."],
[/AADSTS50011|redirect_uri/i,"La URL de redirección no coincide con la registrada en Azure AD."],
[/AADSTS700054|AADSTS9002326|cross-origin/i,"La URL de redirección debe estar registrada como Aplicación de página única (SPA)."],
[/popup_window_error|empty_window_error|user_cancelled|monitor_window_timeout/i,"El navegador bloqueó o cerró la ventana de inicio de sesión."],
[/crypto_nonexistent|insecure/i,"MSAL requiere https:// o http://localhost."],
[/401|InvalidAuthenticationToken/i,"Graph rechazó el token. Cierra sesión, recarga y vuelve a iniciar sesión."],
[/403|accessDenied|Access denied/i,"El usuario no tiene permisos sobre SharePoint."],
[/No se encontró ".*equipos/i,"No se encontró equipos.json en la ruta configurada."],
[/no es un JSON válido|no contiene un arreglo/i,"equipos.json no tiene el formato esperado."],
[/404|itemNotFound|Invalid hostname/i,"No se encontró el sitio, lista o archivo."],
[/Failed to fetch|NetworkError|AbortError/i,"No hubo respuesta de Microsoft Graph."]
];
const hallada=ayudas.find(a=>a[0].test(t));
return hallada?hallada[1]:"";
}

const peticionesFallidas=[];

function registrarFallo(status,url){
let corta=String(url).replace("https://graph.microsoft.com/v1.0","");
try{corta=decodeURIComponent(corta);}catch(e){}
const linea=status+"  "+corta.slice(0,170);
if(!peticionesFallidas.includes(linea)){peticionesFallidas.push(linea);}
console.warn("[Graph] petición fallida:",status,corta);
}

const diagnosticoLineas=[];

function mostrarDiagnostico(paso,error){
try{
const mensaje=String(((error&&(error.errorCode?error.errorCode+": ":""))+(error&&(error.errorMessage||error.message)))||error||"").slice(0,700);
diagnosticoLineas.push({paso:paso,mensaje:mensaje,ayuda:ayudaError(mensaje+" "+(error&&error.errorCode||""))});

let panel=document.getElementById("panelDiagnostico");
if(!panel){
panel=document.createElement("div");
panel.id="panelDiagnostico";
panel.style.cssText="position:fixed;right:16px;bottom:16px;z-index:99999;max-width:520px;max-height:70vh;overflow:auto;background:#fff;color:#172033;border:2px solid #c0392b;border-radius:12px;box-shadow:0 10px 40px rgba(0,0,0,.35);padding:14px 16px;font:12px/1.5 Arial,sans-serif;";
document.body.appendChild(panel);
}

panel.textContent="";

const titulo=document.createElement("strong");
titulo.textContent="Diagnóstico de carga";
titulo.style.cssText="display:block;margin-bottom:6px;color:#c0392b;font-size:13px;";
panel.appendChild(titulo);

const contexto=document.createElement("div");
contexto.style.cssText="color:#667085;margin-bottom:8px;";
contexto.textContent="Página: "+location.origin+" | MSAL: "+(msalInstanceInfraestructura?"cargado":"NO cargó")+" | Sesiones: "+(msalInstanceInfraestructura?msalInstanceInfraestructura.getAllAccounts().length:0)+" | Token: "+(tokenActual?"sí":"no");
panel.appendChild(contexto);

if(peticionesFallidas.length){
const pf=document.createElement("div");
pf.style.cssText="background:#fdecea;border:1px solid #f1a9a0;border-radius:8px;padding:8px;margin-bottom:8px;";
const pt=document.createElement("strong");
pt.textContent="Peticiones a Microsoft Graph que fallaron:";
pf.appendChild(pt);
peticionesFallidas.slice(-8).forEach(l=>{
const d=document.createElement("div");
d.textContent=l;
d.style.cssText="font-family:Consolas,monospace;font-size:11px;word-break:break-all;";
pf.appendChild(d);
});
panel.appendChild(pf);
}

diagnosticoLineas.forEach(l=>{
const bloque=document.createElement("div");
bloque.style.cssText="border-top:1px solid #e4e7ec;padding:7px 0;";
const p=document.createElement("strong");
p.textContent=l.paso;
const m=document.createElement("div");
m.textContent=l.mensaje;
m.style.cssText="font-family:Consolas,monospace;font-size:11px;word-break:break-word;color:#344054;";
bloque.appendChild(p);
bloque.appendChild(m);
if(l.ayuda){
const a=document.createElement("div");
a.textContent="→ "+l.ayuda;
a.style.cssText="margin-top:3px;color:#0a4da2;font-weight:700;";
bloque.appendChild(a);
}
panel.appendChild(bloque);
});

const cerrar=document.createElement("button");
cerrar.type="button";
cerrar.textContent="Cerrar";
cerrar.style.cssText="margin-top:8px;padding:5px 12px;border:1px solid #cbd5e1;border-radius:7px;background:#f8fafc;cursor:pointer;";
cerrar.addEventListener("click",()=>panel.remove());
panel.appendChild(cerrar);
}catch(e){
console.error("No se pudo mostrar el diagnóstico:",e);
}
}

window.addEventListener("unhandledrejection",e=>mostrarDiagnostico("Error no controlado",e.reason));
window.addEventListener("error",e=>mostrarDiagnostico("Error de script",e.error||e.message));

const SCOPES_GRAPH=["User.Read","Sites.Read.All","Sites.ReadWrite.All"];
const CACHE_PREFIJO="alferza_ti_";

let etagEquipos=null;
let promesaSitio=null;
let promesaDrive=null;
let promesaTrabajadores=null;
let trabajadoresListos=false;
let cargandoEquipos=false;
let ultimaActualizacion=0;

function leerCache(clave){
try{
const valor=localStorage.getItem(CACHE_PREFIJO+clave);
return valor?JSON.parse(valor):null;
}catch(e){
return null;
}
}

function guardarCache(clave,valor){
try{
localStorage.setItem(CACHE_PREFIJO+clave,JSON.stringify(valor));
}catch(e){
console.warn("No se pudo guardar en caché:",clave);
}
}

async function obtenerTokenSilencioso(){
if(!msalInstanceInfraestructura){return null;}

const cuentas=msalInstanceInfraestructura.getAllAccounts();
if(!cuentas.length){return null;}

cuentaActual=cuentas[0];

try{
const r=await msalInstanceInfraestructura.acquireTokenSilent({scopes:SCOPES_GRAPH,account:cuentaActual});
tokenActual=r.accessToken;
return tokenActual;
}catch(error){
console.warn("Token silencioso no disponible:",error);
mostrarDiagnostico("Token silencioso",error);
return null;
}
}

async function iniciarSesion(interactivo){
const silencioso=await obtenerTokenSilencioso();
if(silencioso||!interactivo){return silencioso;}

const r=cuentaActual
?await msalInstanceInfraestructura.acquireTokenPopup({scopes:SCOPES_GRAPH,account:cuentaActual})
:await msalInstanceInfraestructura.loginPopup({scopes:SCOPES_GRAPH});

cuentaActual=r.account;
tokenActual=r.accessToken||(await obtenerTokenSilencioso());
return tokenActual;
}

async function graphFetch(url,opciones,intento){
intento=intento||0;

const {timeoutMs=30000,...resto}=opciones||{};
const ctrl=new AbortController();
const timer=setTimeout(()=>ctrl.abort(),timeoutMs);

try{
const response=await fetch(url,{...resto,signal:ctrl.signal,headers:{Authorization:`Bearer ${tokenActual}`,...(resto.headers||{})}});

if(!response.ok){
registrarFallo(response.status,url);
}

if((response.status===429||response.status===503)&&intento<3){
const espera=(parseInt(response.headers.get("Retry-After"),10)||2**intento)*1000;
await new Promise(r=>setTimeout(r,Math.min(espera,10000)));
return graphFetch(url,opciones,intento+1);
}

if(response.status===401&&intento<1&&(await obtenerTokenSilencioso())){
return graphFetch(url,opciones,intento+1);
}

return response;
}finally{
clearTimeout(timer);
}
}

function obtenerSitioSharePoint(){
if(!promesaSitio){
promesaSitio=(async()=>{
const clave="sitio:"+SHAREPOINT_SITE;
const guardado=leerCache(clave);
if(guardado&&guardado.id){return guardado;}

const response=await graphFetch(`https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOSTNAME}:${SHAREPOINT_SITE}?$select=id`);
if(!response.ok){
throw new Error(`No se pudo obtener el sitio SharePoint. ${response.status} ${await response.text()}`);
}

const sitio=await response.json();
guardarCache(clave,{id:sitio.id});
return {id:sitio.id};
})().catch(error=>{
promesaSitio=null;
throw error;
});
}
return promesaSitio;
}

function obtenerDriveSharePoint() {
    if (!promesaDrive) {
        promesaDrive = Promise.resolve({
            id: "b!1l0EwOKXpU24cUGYaNjPRiTQWIyIIChCkDDL197LcoOFOD7ThItqTIJ_Z1KoYrrS"
        });
    }

    return promesaDrive;
}


let ubicacionEquipos=null;

async function buscarEquiposJson(raiz){
const nombre=RUTA_EQUIPOS.split("/").pop();

const r=await graphFetch(`${raiz}/root/search(q='${encodeURIComponent(nombre)}')?$select=id,name,eTag,parentReference&$top=25`);

if(r.ok){
const items=(await r.json()).value||[];
const encontrados=items.filter(i=>normalizarTexto(i.name)===normalizarTexto(nombre));

if(encontrados.length>0){
const it=encontrados[0];
return {
meta:`${raiz}/items/${it.id}`,
contenido:`${raiz}/items/${it.id}/content`,
etag:it.eTag||null
};
}
}

throw new Error(`No se encontró "${RUTA_EQUIPOS}" en el OneDrive/sitio configurado.`);
}

async function leerJsonEquipos(response){
const buffer=await response.arrayBuffer();
const b=new Uint8Array(buffer);

let codificacion="utf-8";
if(b[0]===0xFF&&b[1]===0xFE){
codificacion="utf-16le";
}else if(b[0]===0xFE&&b[1]===0xFF){
codificacion="utf-16be";
}

const texto=new TextDecoder(codificacion).decode(buffer).replace(/^\uFEFF/,"");

let datos;
try{
datos=JSON.parse(texto);
}catch(e){
throw new Error(`equipos.json no es un JSON válido: ${e.message}`);
}

if(Array.isArray(datos)){return datos;}

if(datos&&typeof datos==="object"){
for(const clave of ["value","equipos","Equipos","data","items"]){
if(Array.isArray(datos[clave])){return datos[clave];}
}
if(datos.Nombre){return [datos];}
}

throw new Error("equipos.json no contiene un arreglo de equipos.");
}

async function obtenerEquiposDesdeSharePoint(){
const drive=await obtenerDriveSharePoint();
const raiz=`https://graph.microsoft.com/v1.0/drives/${drive.id}`;
const ruta=RUTA_EQUIPOS.split("/").map(p=>encodeURIComponent(p)).join("/");

const candidatos=[];
if(ubicacionEquipos){candidatos.push(ubicacionEquipos);}
candidatos.push({
meta:`${raiz}/root:/${ruta}`,
contenido:`${raiz}/root:/${ruta}:/content`
});

let ubicacion=null;
let etag=null;

for(const c of candidatos){
const r=await graphFetch(`${c.meta}?$select=eTag`);
if(r.ok){
etag=(await r.json()).eTag||null;
ubicacion=c;
break;
}
if(r.status!==404){
throw new Error(`No se pudo consultar equipos.json. ${r.status} ${await r.text()}`);
}
}

if(!ubicacion){
ubicacion=await buscarEquiposJson(raiz);
etag=ubicacion.etag;
}

ubicacionEquipos=ubicacion;

if(etag&&etag===etagEquipos&&equiposData.length>0){
return null;
}

const response=await graphFetch(ubicacion.contenido,{timeoutMs:90000});
if(!response.ok){
throw new Error(`No se pudo descargar equipos.json. ${response.status} ${await response.text()}`);
}

const datos=await leerJsonEquipos(response);
etagEquipos=etag;
return datos;
}

async function obtenerListaTrabajadores(){
if(listaTrabajadores){return listaTrabajadores;}

const sitio=await obtenerSitioSharePoint();
const response=await graphFetch(`https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists`);

if(!response.ok){
throw new Error(`No se pudieron obtener las listas SharePoint. ${response.status} ${await response.text()}`);
}

const data=await response.json();
const listas=data.value||[];

listaTrabajadores=listas.find(lista=>normalizarTexto(lista.displayName)===normalizarTexto(SHAREPOINT_LIST_TRABAJADORES));

if(!listaTrabajadores){
throw new Error(`No se encontró la lista "${SHAREPOINT_LIST_TRABAJADORES}".`);
}

return listaTrabajadores;
}

async function obtenerColumnasTrabajadores(){
if(columnasTrabajadores){return columnasTrabajadores;}

const lista=await obtenerListaTrabajadores();
const sitio=await obtenerSitioSharePoint();
const response=await graphFetch(`https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/columns`);

if(!response.ok){
throw new Error(`No se pudieron obtener las columnas de TrabajadoresEquipos. ${response.status} ${await response.text()}`);
}

const data=await response.json();
columnasTrabajadores=data.value||[];
return columnasTrabajadores;
}

function obtenerNombreInternoColumna(columnas,nombreVisible){
const columna=columnas.find(item=>normalizarTexto(item.displayName)===normalizarTexto(nombreVisible));
return columna?columna.name:null;
}

function obtenerTrabajadores(){
if(trabajadoresCargados){return Promise.resolve(trabajadoresPorEquipo);}

if(!promesaTrabajadores){
promesaTrabajadores=cargarTrabajadoresDesdeSharePoint().finally(()=>{promesaTrabajadores=null;});
}

return promesaTrabajadores;
}

async function cargarTrabajadoresDesdeSharePoint(){
try{
const sitio=await obtenerSitioSharePoint();
const lista=await obtenerListaTrabajadores();

let url=`https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items`+`?expand=fields($select=Equipo1,NombreTrabajador)&$top=999`;
const items=[];

while(url){
const response=await graphFetch(url);
if(!response.ok){
throw new Error(`Error cargando trabajadores: ${response.status} ${await response.text()}`);
}
const data=await response.json();
items.push(...(data.value||[]));
url=data["@odata.nextLink"]||null;
}

trabajadoresData=items;
trabajadoresPorEquipo={};

items.forEach(item=>{
const fields=item.fields||{};
if(fields.Equipo1&&fields.NombreTrabajador){
trabajadoresPorEquipo[normalizarTexto(fields.Equipo1)]={
nombre:fields.NombreTrabajador,
itemId:item.id,
equipo:fields.Equipo1
};
}
});

console.log("Asignaciones de trabajadores cargadas:",Object.keys(trabajadoresPorEquipo).length);

if(items.length>0&&Object.keys(trabajadoresPorEquipo).length===0){
await cargarTrabajadoresConColumnas();
}

trabajadoresCargados=true;
trabajadoresListos=true;

if(equiposData.length>0){renderizarTablaEquipos();}

return trabajadoresPorEquipo;
}catch(error){
console.warn("No se pudieron cargar las asignaciones:",error);
trabajadoresListos=true;
if(equiposData.length>0){renderizarTablaEquipos();}
return trabajadoresPorEquipo;
}
}

async function cargarTrabajadoresConColumnas(){
const columnas=await obtenerColumnasTrabajadores();
const campoEquipo=obtenerNombreInternoColumna(columnas,"Equipo1");
const campoTrabajador=obtenerNombreInternoColumna(columnas,"Trabajador");

if(!campoEquipo||!campoTrabajador){return;}

trabajadoresPorEquipo={};

trabajadoresData.forEach(item=>{
const fields=item.fields||{};
const equipo=fields[campoEquipo];
const trabajador=fields[campoTrabajador];
if(equipo&&trabajador){
trabajadoresPorEquipo[normalizarTexto(equipo)]={nombre:trabajador,itemId:item.id,equipo:equipo};
}
});
}

function obtenerTrabajadorEquipo(nombreEquipo){
return trabajadoresPorEquipo[normalizarTexto(nombreEquipo)]||null;
}

async function obtenerMonitoreoTI(){
const sitio=await obtenerSitioSharePoint();
const url=`https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${encodeURIComponent(SHAREPOINT_LIST_SERVICIOS)}/items?expand=fields&$top=999`;
const response=await graphFetch(url);

if(!response.ok){
throw new Error(`No se pudo obtener MonitoreoTI. ${response.status} ${await response.text()}`);
}

return (await response.json()).value||[];
}

function obtenerCampo(fields,nombres,defecto){
const mapa={};
Object.keys(fields).forEach(clave=>{mapa[sinTildes(clave)]=fields[clave];});

for(const nombre of nombres){
const valor=mapa[sinTildes(nombre)];
if(valor!==undefined&&valor!==null&&String(valor).trim()!==""){return valor;}
}

return defecto;
}

function adivinarDescripcion(fields,excluir){
const claves=excluir.map(sinTildes);
let mejor="";

Object.entries(fields).forEach(([clave,valor])=>{
if(typeof valor!=="string"){return;}
if(clave.startsWith("@")||clave.startsWith("_")){return;}
if(claves.includes(sinTildes(valor))){return;}
if(/^\d{4}-\d{2}-\d{2}T/.test(valor)){return;}
if(valor.length>mejor.length){mejor=valor;}
});

return mejor.length>=15?mejor:"";
}

function formatearRevision(valor){
if(!valor){return "--";}

const texto=String(valor).trim();

if(/^\d{1,2}\/\d{1,2}\/\d{4}/.test(texto)){return texto;}

const fecha=new Date(texto);
if(isNaN(fecha.getTime())){return texto;}

return fecha.toLocaleString("es-PE",{timeZone:"America/Lima",day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit",hour12:false});
}

function renderizarServicios(items){
if(!serviciosGrid){return;}

if(!items||items.length===0){
serviciosGrid.innerHTML=`<div class="service-card service-loading"><span>No hay servicios registrados.</span></div>`;
estadoGeneral.textContent="● SIN DATOS";
estadoGeneral.className="section-status warning";
return;
}

let serviciosOnline=0;

serviciosGrid.innerHTML=items.map(item=>{
const fields=item.fields||{};

const nombre=obtenerCampo(fields,["Servicio","Nombre","Title","LinkTitle"],"Servicio");
const estado=obtenerCampo(fields,["Estado","Status"],"Sin estado");
const descripcion=obtenerCampo(fields,["Descripcion","Description","Detalle","Mensaje","Observacion","Comentario"],"")||adivinarDescripcion(fields,[nombre,estado])||"Sin descripción";
const latencia=obtenerCampo(fields,["Latencia","Ping","Tiempo"],"--");
const revisionRaw=obtenerCampo(fields,["UltimaRevision","Ultima revision","Ultima_revision","Fecha","FechaRevision"],null);
const revision=formatearRevision(revisionRaw);

const estadoNormalizado=sinTildes(estado);
let claseEstado="neutral";

if(estadoNormalizado.includes("operativo")||estadoNormalizado.includes("online")||estadoNormalizado.includes("conectado")){
claseEstado="online";
serviciosOnline++;
}else if(estadoNormalizado.includes("mantenimiento")||estadoNormalizado.includes("advertencia")||estadoNormalizado.includes("warning")||estadoNormalizado.includes("lento")){
claseEstado="warning";
}else if(estadoNormalizado.includes("caido")||estadoNormalizado.includes("offline")||estadoNormalizado.includes("desconectado")||estadoNormalizado.includes("error")){
claseEstado="offline";
}

return `
<article class="service-card ${claseEstado}">
<div class="service-main">
<div class="service-icon">🖥️</div>
<div class="service-info">
<div class="service-heading">
<h3>${escaparHTML(nombre)}</h3>
<span class="service-status ${claseEstado}"><span class="status-dot"></span>${escaparHTML(estado)}</span>
</div>
<p>${escaparHTML(descripcion)}</p>
</div>
</div>
<div class="service-meta">
<div class="service-latency"><span>Latencia</span><strong>${escaparHTML(latencia)}</strong></div>
<div class="service-revision">
<span class="revision-icon">🕐</span>
<div><span>Última revisión</span><strong>${escaparHTML(revision)}</strong></div>
</div>
</div>
</article>
`;
}).join("");

if(serviciosOnline===items.length){
estadoGeneral.textContent="● OPERATIVO";
estadoGeneral.className="section-status online";
}else if(serviciosOnline>0){
estadoGeneral.textContent="● REVISAR";
estadoGeneral.className="section-status warning";
}else{
estadoGeneral.textContent="● NO DISPONIBLE";
estadoGeneral.className="section-status offline";
}
}

function actualizarResumenEquipos(){
const total=equiposData.length;
const conectados=equiposData.filter(e=>normalizarTexto(e.Estado)==="conectado").length;
const desconectados=equiposData.filter(e=>normalizarTexto(e.Estado)==="desconectado").length;

totalEquipos.textContent=total;
equiposConectados.textContent=conectados;
equiposDesconectados.textContent=desconectados;
}

function actualizarFiltroAreas(){
const areas=[...new Set(equiposData.map(e=>e.Area).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),"es"));

filtroAreaEquipo.innerHTML=`<option value="todos">Todas las áreas</option>`;

areas.forEach(area=>{
const option=document.createElement("option");
option.value=area;
option.textContent=area;
filtroAreaEquipo.appendChild(option);
});
}

function aplicarFiltrosEquipos(conservarPagina){
const texto=normalizarTexto(buscarEquipo.value);
const estado=filtroEstadoEquipo.value;
const area=filtroAreaEquipo.value;

equiposFiltrados=equiposData.filter(equipo=>{
const coincideTexto=!texto||normalizarTexto(equipo.Nombre).includes(texto)||normalizarTexto(equipo.IP).includes(texto)||normalizarTexto(equipo.DNSHostName).includes(texto);
const coincideEstado=estado==="todos"||String(equipo.Estado||"")===estado;
const coincideArea=area==="todos"||String(equipo.Area||"")===area;
return coincideTexto&&coincideEstado&&coincideArea;
});

if(conservarPagina!==true){paginaActual=1;}

renderizarTablaEquipos();
}

function obtenerClaseEstado(estado){
const e=normalizarTexto(estado);
if(e==="conectado"){return "online";}
if(e==="desconectado"){return "offline";}
return "neutral";
}

function renderizarTablaEquipos(){
const totalResultados=equiposFiltrados.length;
const totalPaginas=Math.max(1,Math.ceil(totalResultados/EQUIPOS_POR_PAGINA));

if(paginaActual>totalPaginas){paginaActual=totalPaginas;}

const inicio=(paginaActual-1)*EQUIPOS_POR_PAGINA;
const fin=inicio+EQUIPOS_POR_PAGINA;
const equiposPagina=equiposFiltrados.slice(inicio,fin);

if(equiposResultados){
equiposResultados.textContent=totalResultados===0?"0 equipos encontrados":`Equipos encontrados: ${totalResultados}`;
}

if(totalResultados===0){
equiposTableBody.innerHTML="";
equiposEmpty.style.display="flex";
renderizarPaginacion(0);
return;
}

equiposEmpty.style.display="none";

equiposTableBody.innerHTML=equiposPagina.map(equipo=>{
const estado=equipo.Estado||"No verificable";
const claseEstado=obtenerClaseEstado(estado);
const trabajador=obtenerTrabajadorEquipo(equipo.Nombre);

const trabajadorHTML=trabajador
?`<span class="trabajador-name" title="${escaparHTML(trabajador.nombre)}">${escaparHTML(trabajador.nombre)}</span>`
:!trabajadoresListos
?`<span style="color:#98a2b3;font-size:12px;">Cargando…</span>`
:`<span style="color:#98a2b3;font-size:12px;">Sin asignar</span>`;

return `
<tr>
<td>
<div class="equipo-name">
<div class="equipo-icon">🖥️</div>
<div>
<strong title="${escaparHTML(equipo.Nombre)}">${escaparHTML(equipo.Nombre)}</strong>
<span title="${escaparHTML(equipo.DNSHostName)}">${escaparHTML(equipo.DNSHostName||"Sin DNS")}</span>
</div>
</div>
</td>
<td><span class="equipo-status ${claseEstado}"><span class="equipo-status-dot"></span>${escaparHTML(estado)}</span></td>
<td class="trabajador-cell">${trabajadorHTML}</td>
<td><span class="equipo-ip">${escaparHTML(equipo.IP||"--")}</span></td>
<td>
<div class="equipo-so">
${escaparHTML(equipo.SistemaOperativo||"--")}
${equipo.VersionSO?`<br><small>${escaparHTML(equipo.VersionSO)}</small>`:""}
</div>
</td>
<td><span class="equipo-area" title="${escaparHTML(equipo.Area)}">${escaparHTML(equipo.Area||"--")}</span></td>
<td><span class="equipo-date">${escaparHTML(equipo.UltimoRegistroAD||"--")}</span></td>
</tr>
`;
}).join("");

renderizarPaginacion(totalPaginas);
}

function renderizarPaginacion(totalPaginas){
const existente=document.querySelector(".equipos-pagination");
if(existente){existente.remove();}

if(totalPaginas<=1){return;}

const contenedor=document.createElement("div");
contenedor.className="equipos-pagination";

const botonAnterior=document.createElement("button");
botonAnterior.type="button";
botonAnterior.className="pagination-btn pagination-prev";
botonAnterior.innerHTML="‹";
botonAnterior.disabled=paginaActual===1;
botonAnterior.addEventListener("click",()=>{
if(paginaActual>1){
paginaActual--;
renderizarTablaEquipos();
}
});
contenedor.appendChild(botonAnterior);

generarPaginas(paginaActual,totalPaginas).forEach(pagina=>{
if(pagina==="..."){
const puntos=document.createElement("span");
puntos.className="pagination-ellipsis";
puntos.textContent="...";
contenedor.appendChild(puntos);
return;
}

const boton=document.createElement("button");
boton.type="button";
boton.className="pagination-btn";
boton.textContent=pagina;
if(pagina===paginaActual){boton.classList.add("active");}
boton.addEventListener("click",()=>{
paginaActual=pagina;
renderizarTablaEquipos();
});
contenedor.appendChild(boton);
});

const botonSiguiente=document.createElement("button");
botonSiguiente.type="button";
botonSiguiente.className="pagination-btn pagination-next";
botonSiguiente.innerHTML="›";
botonSiguiente.disabled=paginaActual===totalPaginas;
botonSiguiente.addEventListener("click",()=>{
if(paginaActual<totalPaginas){
paginaActual++;
renderizarTablaEquipos();
}
});
contenedor.appendChild(botonSiguiente);

const tableContainer=document.querySelector(".equipos-table-container");
if(tableContainer){
tableContainer.parentNode.insertBefore(contenedor,tableContainer.nextSibling);
}
}

function generarPaginas(actual,total){
if(total<=7){
return Array.from({length:total},(_,i)=>i+1);
}

const paginas=[1];
if(actual>4){paginas.push("...");}

const inicio=Math.max(2,actual-1);
const fin=Math.min(total-1,actual+1);

for(let i=inicio;i<=fin;i++){paginas.push(i);}

if(actual<total-3){paginas.push("...");}

paginas.push(total);
return paginas;
}

let equipoGestionSeleccionado=null;

function abrirGestionTrabajadores(){
if(!gestionTrabajadorModal){return;}

equipoGestionSeleccionado=null;
gestionIP.value="";
gestionTrabajador.value="";
gestionTrabajador.disabled=true;
equipoEncontrado.style.display="none";
gestionTrabajadorMensaje.style.display="none";
gestionTrabajadorMensaje.textContent="";
btnRegistrarTrabajador.disabled=true;
btnEditarTrabajador.disabled=true;
btnEliminarTrabajador.disabled=true;

gestionTrabajadorModal.style.display="flex";

setTimeout(()=>gestionIP.focus(),100);
}

function cerrarGestionTrabajadores(){
gestionTrabajadorModal.style.display="none";
equipoGestionSeleccionado=null;
gestionIP.value="";
gestionTrabajador.value="";
gestionTrabajador.disabled=true;
equipoEncontrado.style.display="none";
gestionTrabajadorMensaje.style.display="none";
gestionTrabajadorMensaje.textContent="";
btnRegistrarTrabajador.disabled=true;
btnEditarTrabajador.disabled=true;
btnEliminarTrabajador.disabled=true;
}

function mostrarMensajeGestion(mensaje,tipo="error"){
gestionTrabajadorMensaje.textContent=mensaje;
gestionTrabajadorMensaje.style.display="block";
gestionTrabajadorMensaje.style.color=tipo==="success"?"#027a48":"#b42318";
}

async function buscarEquipoPorIP(){
const ip=gestionIP.value.trim();

if(!ip){
mostrarMensajeGestion("Ingresa una dirección IP.");
return;
}

gestionTrabajadorMensaje.style.display="none";
equipoEncontrado.style.display="none";
gestionTrabajador.disabled=true;
btnRegistrarTrabajador.disabled=true;
btnEditarTrabajador.disabled=true;
btnEliminarTrabajador.disabled=true;
equipoGestionSeleccionado=null;

buscarEquipoIP.disabled=true;
buscarEquipoIP.textContent="Buscando...";

try{
if(!equiposData||equiposData.length===0){
await cargarEquipos();
}

const equipo=equiposData.find(item=>normalizarTexto(item.IP)===normalizarTexto(ip));

if(!equipo){
mostrarMensajeGestion(`No se encontró ningún equipo con la IP ${ip}.`);
return;
}

equipoGestionSeleccionado=equipo;

gestionEquipoNombre.textContent=equipo.Nombre||"Sin nombre";
gestionEquipoDetalle.textContent=`${equipo.IP||ip}`+` · ${equipo.DNSHostName||"Sin DNS"}`+` · ${equipo.Area||"Sin área"}`;

equipoEncontrado.style.display="flex";
gestionTrabajador.disabled=false;

let trabajador=obtenerTrabajadorEquipo(equipo.Nombre);

if(!trabajadoresCargados){
try{
await obtenerTrabajadores();
trabajador=obtenerTrabajadorEquipo(equipo.Nombre);
}catch(error){
console.warn("No se pudo consultar la asignación:",error);
}
}

if(trabajador){
gestionTrabajador.value=trabajador.nombre;
btnRegistrarTrabajador.disabled=true;
btnEditarTrabajador.disabled=false;
btnEliminarTrabajador.disabled=false;
}else{
gestionTrabajador.value="";
btnRegistrarTrabajador.disabled=false;
btnEditarTrabajador.disabled=true;
btnEliminarTrabajador.disabled=true;
}

gestionTrabajador.focus();
}catch(error){
console.error("Error buscando equipo por IP:",error);
mostrarMensajeGestion(obtenerMensajeError(error));
}finally{
buscarEquipoIP.disabled=false;
buscarEquipoIP.textContent="Buscar";
}
}

async function registrarTrabajadorPorIP(){
if(!equipoGestionSeleccionado){
mostrarMensajeGestion("Primero busca un equipo por IP.");
return;
}

const nombre=gestionTrabajador.value.trim();

if(!nombre){
mostrarMensajeGestion("Ingresa el nombre del trabajador.");
gestionTrabajador.focus();
return;
}

const equipo=equipoGestionSeleccionado;

let asignacionActual=obtenerTrabajadorEquipo(equipo.Nombre);

if(!asignacionActual&&!trabajadoresCargados){
await obtenerTrabajadores();
asignacionActual=obtenerTrabajadorEquipo(equipo.Nombre);
}

if(asignacionActual){
mostrarMensajeGestion(`Este equipo ya está asignado a ${asignacionActual.nombre}.`);
return;
}

try{
btnRegistrarTrabajador.disabled=true;
btnRegistrarTrabajador.textContent="Registrando...";

const sitio=await obtenerSitioSharePoint();
const lista=await obtenerListaTrabajadores();
const columnas=await obtenerColumnasTrabajadores();

const campoEquipo=obtenerNombreInternoColumna(columnas,"Equipo1");
const campoTrabajador=obtenerNombreInternoColumna(columnas,"NombreTrabajador");

if(!campoEquipo){throw new Error('No se encontró la columna "Equipo1".');}
if(!campoTrabajador){throw new Error('No se encontró la columna "NombreTrabajador".');}
if(campoTrabajador==="Title"||campoTrabajador==="LinkTitle"){
throw new Error(`La columna "NombreTrabajador" está resolviendo incorrectamente a "${campoTrabajador}".`);
}

const fields={};
fields[campoEquipo]=equipo.Nombre;
fields[campoTrabajador]=nombre;

const response=await graphFetch(`https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items`,{
method:"POST",
headers:{"Content-Type":"application/json"},
body:JSON.stringify({fields})
});

if(!response.ok){
const texto=await response.text();
throw new Error(`SharePoint respondió ${response.status}: ${texto}`);
}

const resultado=await response.json();

trabajadoresPorEquipo[normalizarTexto(equipo.Nombre)]={nombre:nombre,itemId:resultado.id,equipo:equipo.Nombre};

trabajadoresCargados=true;
trabajadoresListos=true;

renderizarTablaEquipos();

gestionTrabajador.value=nombre;
btnRegistrarTrabajador.disabled=true;
btnEditarTrabajador.disabled=false;
btnEliminarTrabajador.disabled=false;

mostrarMensajeGestion("Trabajador registrado correctamente.","success");
}catch(error){
console.error("Error registrando trabajador:",error);
mostrarMensajeGestion(obtenerMensajeError(error));
}finally{
btnRegistrarTrabajador.textContent="Registrar";
if(equipoGestionSeleccionado&&!obtenerTrabajadorEquipo(equipoGestionSeleccionado.Nombre)){
btnRegistrarTrabajador.disabled=false;
}
}
}

async function editarTrabajadorPorIP(){
if(!equipoGestionSeleccionado){
mostrarMensajeGestion("Primero busca un equipo por IP.");
return;
}

const nombre=gestionTrabajador.value.trim();

if(!nombre){
mostrarMensajeGestion("Ingresa el nombre del trabajador.");
gestionTrabajador.focus();
return;
}

const equipo=equipoGestionSeleccionado;
const asignacion=obtenerTrabajadorEquipo(equipo.Nombre);

if(!asignacion||!asignacion.itemId){
mostrarMensajeGestion("Este equipo no tiene un trabajador registrado.");
return;
}

try{
btnEditarTrabajador.disabled=true;
btnEditarTrabajador.textContent="Guardando...";

const sitio=await obtenerSitioSharePoint();
const lista=await obtenerListaTrabajadores();
const columnas=await obtenerColumnasTrabajadores();

const campoTrabajador=obtenerNombreInternoColumna(columnas,"NombreTrabajador");

if(!campoTrabajador){throw new Error('No se encontró la columna "NombreTrabajador".');}

const fields={};
fields[campoTrabajador]=nombre;

const response=await graphFetch(`https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items/${asignacion.itemId}/fields`,{
method:"PATCH",
headers:{"Content-Type":"application/json"},
body:JSON.stringify(fields)
});

if(!response.ok){
const texto=await response.text();
throw new Error(`SharePoint respondió ${response.status}: ${texto}`);
}

trabajadoresPorEquipo[normalizarTexto(equipo.Nombre)].nombre=nombre;

renderizarTablaEquipos();

gestionTrabajador.value=nombre;
mostrarMensajeGestion("Trabajador actualizado correctamente.","success");
}catch(error){
console.error("Error editando trabajador:",error);
mostrarMensajeGestion(obtenerMensajeError(error));
}finally{
btnEditarTrabajador.disabled=false;
btnEditarTrabajador.textContent="Editar";
}
}

async function eliminarTrabajadorPorIP(){
if(!equipoGestionSeleccionado){
mostrarMensajeGestion("Primero busca un equipo por IP.");
return;
}

const equipo=equipoGestionSeleccionado;
const asignacion=obtenerTrabajadorEquipo(equipo.Nombre);

if(!asignacion||!asignacion.itemId){
mostrarMensajeGestion("Este equipo no tiene un trabajador registrado.");
return;
}

const confirmar=confirm(`¿Deseas eliminar la asignación de ${asignacion.nombre} del equipo ${equipo.Nombre}?`);
if(!confirmar){return;}

try{
btnEliminarTrabajador.disabled=true;
btnEliminarTrabajador.textContent="Eliminando...";

const sitio=await obtenerSitioSharePoint();
const lista=await obtenerListaTrabajadores();

const response=await graphFetch(`https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items/${asignacion.itemId}`,{method:"DELETE"});

if(!response.ok){
const texto=await response.text();
throw new Error(`SharePoint respondió ${response.status}: ${texto}`);
}

delete trabajadoresPorEquipo[normalizarTexto(equipo.Nombre)];

renderizarTablaEquipos();

gestionTrabajador.value="";
btnRegistrarTrabajador.disabled=false;
btnEditarTrabajador.disabled=true;
btnEliminarTrabajador.disabled=true;

mostrarMensajeGestion("Asignación eliminada correctamente.","success");
}catch(error){
console.error("Error eliminando trabajador:",error);
mostrarMensajeGestion(obtenerMensajeError(error));
}finally{
btnEliminarTrabajador.textContent="Eliminar";
}
}

function obtenerMensajeError(error){
const mensaje=String((error&&error.message)||error||"");

if(mensaje.includes("403")){return "No tienes permisos para modificar esta lista.";}
if(mensaje.includes("401")){return "La sesión expiró. Recarga la página e inicia sesión nuevamente.";}

return mensaje||"No se pudo realizar la operación.";
}

if(btnGestionarTrabajadores){btnGestionarTrabajadores.addEventListener("click",abrirGestionTrabajadores);}
if(buscarEquipoIP){buscarEquipoIP.addEventListener("click",buscarEquipoPorIP);}

if(gestionIP){
gestionIP.addEventListener("keydown",event=>{
if(event.key==="Enter"){
event.preventDefault();
buscarEquipoPorIP();
}
if(event.key==="Escape"){cerrarGestionTrabajadores();}
});
}

if(gestionTrabajador){
gestionTrabajador.addEventListener("keydown",event=>{
if(event.key==="Escape"){cerrarGestionTrabajadores();}
});
}

if(btnRegistrarTrabajador){btnRegistrarTrabajador.addEventListener("click",registrarTrabajadorPorIP);}
if(btnEditarTrabajador){btnEditarTrabajador.addEventListener("click",editarTrabajadorPorIP);}
if(btnEliminarTrabajador){btnEliminarTrabajador.addEventListener("click",eliminarTrabajadorPorIP);}
if(cerrarGestionTrabajadorModal){cerrarGestionTrabajadorModal.addEventListener("click",cerrarGestionTrabajadores);}
if(cancelarGestionTrabajador){cancelarGestionTrabajador.addEventListener("click",cerrarGestionTrabajadores);}

if(gestionTrabajadorModal){
gestionTrabajadorModal.addEventListener("click",event=>{
if(event.target===gestionTrabajadorModal){cerrarGestionTrabajadores();}
});
}

if(buscarEquipo){
let temporizadorBusqueda=null;
buscarEquipo.addEventListener("input",()=>{
clearTimeout(temporizadorBusqueda);
temporizadorBusqueda=setTimeout(()=>aplicarFiltrosEquipos(),200);
});
}

if(filtroEstadoEquipo){filtroEstadoEquipo.addEventListener("change",()=>aplicarFiltrosEquipos());}
if(filtroAreaEquipo){filtroAreaEquipo.addEventListener("change",()=>aplicarFiltrosEquipos());}

function aplicarDatosEquipos(equipos){
equiposData=equipos;

const areaPrevia=filtroAreaEquipo.value;

actualizarResumenEquipos();
actualizarFiltroAreas();

if([...filtroAreaEquipo.options].some(o=>o.value===areaPrevia)){
filtroAreaEquipo.value=areaPrevia;
}

aplicarFiltrosEquipos(true);
}



function usuarioPuedeVerInventario() {
    const permisos = window.alferzaPermisos || {};

    return String(permisos.grupo || "")
        .trim()
        .toUpperCase() === "TI";
}

async function verificarPermisosInventario() {
    if (
        window.alferzaAccessReady &&
        typeof window.alferzaAccessReady.then === "function"
    ) {
        try {
            await window.alferzaAccessReady;
        } catch (error) {
            return false;
        }
    }

    return usuarioPuedeVerInventario();
}


async function cargarEquipos() {
    // El inventario es exclusivo del grupo TI.
    if (!(await verificarPermisosInventario())) {
        return;
    }

    if (cargandoEquipos) {
        return;
    }

    cargandoEquipos = true;

    try {
        // Revalidar antes de consultar los datos.
        if (!(await verificarPermisosInventario())) {
            return;
        }

        equiposEstado.textContent = "● ACTUALIZANDO";
        equiposEstado.className = "section-status warning";

        const equipos = await obtenerEquiposDesdeSharePoint();

        // Evitar mostrar los datos si los permisos cambiaron durante la carga.
        if (!(await verificarPermisosInventario())) {
            equiposData = [];
            equiposFiltrados = [];

            if (equiposTableBody) {
                equiposTableBody.innerHTML = "";
            }

            if (equiposEstado) {
                equiposEstado.textContent = "● RESTRINGIDO";
                equiposEstado.className = "section-status warning";
            }

            return;
        }

        if (equipos) {
            aplicarDatosEquipos(equipos);

            guardarCache("equipos", {
                etag: etagEquipos,
                datos: equipos
            });
        }

        if (lastUpdate) {
            lastUpdate.textContent = obtenerFechaActual();
        }

        equiposEstado.textContent = "● ACTUALIZADO";
        equiposEstado.className = "section-status online";

    } catch (error) {
        console.error("Error cargando equipos:", error);

        // Mostrar diagnósticos solo a usuarios autorizados de TI.
        if (await verificarPermisosInventario()) {
            mostrarDiagnostico("Inventario (equipos.json)", error);

            if (equiposData.length > 0) {
                equiposEstado.textContent = "● SIN ACTUALIZAR";
                equiposEstado.className = "section-status warning";
            } else {
                equiposEstado.textContent = "● ERROR";
                equiposEstado.className = "section-status offline";

                if (equiposTableBody) {
                    equiposTableBody.innerHTML = `
                        <tr>
                            <td colspan="7" class="equipos-loading">
                                <div class="table-loading">
                                    <span>No se pudo cargar el inventario.</span>
                                </div>
                            </td>
                        </tr>
                    `;
                }

                if (equiposResultados) {
                    equiposResultados.textContent =
                        "Error al obtener los equipos.";
                }
            }
        }
    } finally {
        cargandoEquipos = false;
    }
}

async function cargarServicios() {
    try {
        // Los servicios están disponibles para todos los usuarios autorizados.
        const servicios = await obtenerMonitoreoTI();

        renderizarServicios(servicios);

    } catch (error) {
        console.error("Error cargando servicios:", error);

        mostrarDiagnostico(
            "Servicios (lista " + SHAREPOINT_LIST_SERVICIOS + ")",
            error
        );

        if (serviciosGrid) {
            serviciosGrid.innerHTML = `
                <div class="service-card offline">
                    <div class="service-main">
                        <div class="service-icon">⚠️</div>
                        <div class="service-info">
                            <div class="service-heading">
                                <h3>Error de conexión</h3>
                                <span class="service-status offline">
                                    <span class="status-dot"></span>ERROR
                                </span>
                            </div>
                            <p>
                                No se pudieron obtener los servicios desde SharePoint.
                            </p>
                        </div>
                    </div>
                </div>
            `;
        }

        if (estadoGeneral) {
            estadoGeneral.textContent = "● ERROR";
            estadoGeneral.className = "section-status offline";
        }
    }
}

function mostrarBotonLogin(mensaje){
if(loadingOverlay){loadingOverlay.style.display="flex";}
if(loadingMessage){loadingMessage.textContent=mensaje||"";}

if(document.getElementById("btnLoginMs")){return;}

const boton=document.createElement("button");
boton.id="btnLoginMs";
boton.type="button";
boton.className="loading-btn";
boton.textContent="Iniciar sesión con Microsoft";

boton.addEventListener("click",async()=>{
boton.disabled=true;
try{
if(await iniciarSesion(true)){
boton.remove();
ocultarLoading();
iniciarCargas();
}
}catch(error){
console.error("Error de inicio de sesión:",error);
mostrarDiagnostico("Inicio de sesión",error);
loadingMessage.textContent="No se pudo iniciar sesión. Permite las ventanas emergentes e inténtalo de nuevo.";
}finally{
boton.disabled=false;
}
});

const caja=loadingOverlay.querySelector(".loading-content")||loadingOverlay.querySelector(".loading-card");
if(caja){caja.appendChild(boton);}
}

async function iniciarCargas() {
    ultimaActualizacion = Date.now();

    // Los servicios están disponibles para todos los usuarios autorizados.
    cargarServicios();

    // El inventario solo se consulta para el grupo TI.
    if (await verificarPermisosInventario()) {
        cargarEquipos()
            .then(() => {
                if (equiposData.length > 0) {
                    return obtenerTrabajadores();
                }
            })
            .catch(error => {
                console.warn("Carga de inventario falló:", error);
            });
    }
}

async function cargarDashboard(){
try{
if(!msalInstanceInfraestructura){
mostrarLoading("No se pudo cargar MSAL. Revisa tu conexión y recarga.");
return;
}

let conCache=false;

if(msalInstanceInfraestructura.getAllAccounts().length>0){
const cache=leerCache("equipos");
if(cache&&Array.isArray(cache.datos)&&cache.datos.length>0){
etagEquipos=cache.etag||null;
aplicarDatosEquipos(cache.datos);
conCache=true;
}
}

if(!conCache){mostrarLoading("Iniciando sesión...");}

const token=await iniciarSesion(false);

if(!token){
mostrarBotonLogin("Inicia sesión para ver el monitoreo.");
return;
}

ocultarLoading();
iniciarCargas();
}catch(error){
console.error("Error inicializando infraestructura:",error);
mostrarDiagnostico("Inicialización",error);
mostrarBotonLogin("No se pudo cargar el monitoreo.");
}
}

async function actualizarDashboard(){
if(document.hidden||!tokenActual){return;}

try{
if(!(await obtenerTokenSilencioso())){return;}

ultimaActualizacion=Date.now();

await Promise.all([cargarServicios(),cargarEquipos()]);
}catch(error){
console.error("Error en actualización automática:",error);
}
}

function iniciarInfraestructura(){
cargarDashboard();

setTimeout(()=>{
if(equiposData.length===0&&!document.getElementById("panelDiagnostico")){
mostrarDiagnostico("Sin respuesta tras 20 s",new Error("Aún no llegan datos. Si ves el botón «Iniciar sesión con Microsoft», púlsalo."));
}
},20000);

setInterval(actualizarDashboard,INTERVALO_ACTUALIZACION);

document.addEventListener("visibilitychange",()=>{
if(!document.hidden&&Date.now()-ultimaActualizacion>30000){
actualizarDashboard();
}
});
}

if(document.readyState==="loading"){
document.addEventListener("DOMContentLoaded",iniciarInfraestructura);
}else{
iniciarInfraestructura();
}