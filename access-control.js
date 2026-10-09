(function(){
"use strict";

const CONFIG={
clientId:"5d98417c-74a7-4fab-8f2c-41ac127be696",
tenantId:"dbab984f-4bb1-4b60-9dff-da59f54acdf1",
redirectUri:"https://alferzati.github.io/live-office/blank.html",
basePath:"/live-office/",
sharePointHost:"alferzaholding-my.sharepoint.com",
sharePointPath:"/personal/soporte1_alferza_pe",
listName:"PermisosTI",
adminEmails:["soporte1@alferza.pe"]
};

const MODULOS=["Oficina","Personal","Reservas","Salas","Comunicados","Seguridad","Infraestructura","Tickets","Configuracion"];

const PAGINAS={
"index.html":"Oficina",
"personal.html":"Personal",
"reserva.html":"Reservas",
"reservas.html":"Reservas",
"salas.html":"Salas",
"comunicados.html":"Comunicados",
"seguridad.html":"Seguridad",
"infraestructura.html":"Infraestructura",
"tickets.html":"Tickets",
"configuracion.html":"Configuracion"
};

const PAGINAS_PUBLICAS=["login.html","blank.html"];
const CACHE_PERMISOS_MS=5*60*1000;

let instanciaMSAL=null;
let tokenGraph=null;

function normalizar(valor){
return String(valor||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[\s_-]/g,"").toLowerCase().trim();
}

function esVerdadero(valor){
if(valor===true||valor===1){return true;}
return ["true","1","yes","si","sí","verdadero"].includes(String(valor||"").trim().toLowerCase());
}

function leerPermisosCache(correo){
try{
const cache=JSON.parse(sessionStorage.getItem("alferza_permisos")||"null");
if(cache&&cache.correo===correo&&Date.now()-cache.fecha<CACHE_PERMISOS_MS){
return cache.permisos;
}
}catch(e){}
return null;
}

function guardarPermisosCache(correo,permisos){
try{
sessionStorage.setItem("alferza_permisos",JSON.stringify({correo:correo,permisos:permisos,fecha:Date.now()}));
}catch(e){}
}

function mostrarDenegacion(mensaje,volverAlLogin){
function renderizar(){
document.body.innerHTML=`
<main id="alferza-access-denied" style="min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;box-sizing:border-box;background:#f4f6fa;font-family:Arial,sans-serif;color:#172033;text-align:center;">
<section style="width:100%;max-width:460px;padding:36px 28px;background:#fff;border:1px solid #e1e6ef;border-radius:16px;box-sizing:border-box;">
<h1 style="color:#00205c">🔒 Acceso restringido</h1>
<p style="line-height:1.7;color:#596579">${mensaje}</p>
<button id="alferza-access-back" type="button" style="margin-top:16px;padding:12px 20px;border:0;border-radius:8px;background:#00205c;color:white;cursor:pointer;">
${volverAlLogin?"Volver al inicio de sesión":"Volver a Oficina"}
</button>
</section>
</main>
`;

document.getElementById("alferza-access-back").addEventListener("click",function(){
window.location.replace(volverAlLogin?CONFIG.basePath+"login.html":CONFIG.basePath+"index.html");
});
}

if(document.body){
renderizar();
}else{
document.addEventListener("DOMContentLoaded",renderizar,{once:true});
}
}

async function graph(endpoint){
if(!tokenGraph){throw new Error("No existe un token válido de Microsoft Graph.");}

const url=endpoint.startsWith("https://")?endpoint:"https://graph.microsoft.com/v1.0"+endpoint;

const respuesta=await fetch(url,{headers:{Authorization:"Bearer "+tokenGraph,Accept:"application/json"}});

if(!respuesta.ok){
throw new Error("Error de Microsoft Graph: HTTP "+respuesta.status);
}

return respuesta.json();
}

async function obtenerTodos(endpoint){
const resultados=[];
let url=endpoint;

while(url){
const pagina=await graph(url);
resultados.push(...(pagina.value||[]));
url=pagina["@odata.nextLink"]||null;
}

return resultados;
}

async function obtenerListaPermisos(){
const sitio=await graph("/sites/"+CONFIG.sharePointHost+":"+CONFIG.sharePointPath);

const listas=await obtenerTodos("/sites/"+sitio.id+"/lists?$select=id,name,displayName&$top=200");

const lista=listas.find(function(item){
return normalizar(item.displayName)===normalizar(CONFIG.listName)||normalizar(item.name)===normalizar(CONFIG.listName);
});

if(!lista){throw new Error("No se encontró la lista PermisosTI.");}

const registros=await obtenerTodos("/sites/"+sitio.id+"/lists/"+lista.id+"/items?$expand=fields&$top=500");

return registros.map(function(item){return item.fields||{};});
}

async function obtenerPermisos(correo){
const registros=await obtenerListaPermisos();

const registro=registros.find(function(fields){
return String(fields.UsuarioCorreo||"").trim().toLowerCase()===correo;
});

if(!registro){throw new Error("Tu correo no tiene un registro en PermisosTI.");}

const grupo=String(registro.Grupo||"").trim().toLowerCase();

if(grupo==="ti"){
return {correo:correo,grupo:"TI",modulos:[...MODULOS]};
}

if(grupo!=="normal"){
throw new Error("El grupo asignado en PermisosTI no es válido.");
}

const modulos=MODULOS.filter(function(modulo){
const clave=Object.keys(registro).find(function(campo){return normalizar(campo)===normalizar(modulo);});
return clave!==undefined&&esVerdadero(registro[clave]);
});

return {correo:correo,grupo:"Normal",modulos:modulos};
}

function obtenerModuloActual(){
const archivo=window.location.pathname.toLowerCase().split("/").pop();

if(!archivo||archivo==="index.html"){return "Oficina";}
if(archivo==="permisos.html"){return "Permisos";}

return PAGINAS[archivo]||null;
}

function filtrarMenu(permisos){
document.querySelectorAll("a[href]").forEach(function(enlace){
let destino;

try{
destino=new URL(enlace.getAttribute("href"),location.href);
}catch{
return;
}

if(destino.origin!==location.origin){return;}

const archivo=destino.pathname.toLowerCase().split("/").pop();

let modulo=PAGINAS[archivo]||null;

if(archivo===""||archivo==="index.html"){modulo="Oficina";}
if(archivo==="permisos.html"){modulo="Permisos";}

if(!modulo){return;}

const permitido=modulo==="Permisos"?permisos.grupo==="TI":permisos.modulos.includes(modulo);

if(!permitido){
const contenedor=enlace.closest("li")||enlace;
contenedor.style.display="none";
enlace.setAttribute("aria-hidden","true");
enlace.setAttribute("tabindex","-1");
}
});
}

(function filtrarMenuDesdeCache(){
if(sessionStorage.getItem("alferza_login")!=="true"){return;}

let cache=null;
try{cache=JSON.parse(sessionStorage.getItem("alferza_permisos")||"null");}catch(e){}

if(!cache||!cache.permisos||Date.now()-cache.fecha>=CACHE_PERMISOS_MS){return;}

const aplicar=function(){filtrarMenu(cache.permisos);};

if(document.readyState==="loading"){
document.addEventListener("DOMContentLoaded",aplicar,{once:true});
}else{
aplicar();
}
})();

async function validarAcceso(){
const ruta=window.location.pathname;
const archivo=ruta.toLowerCase().split("/").pop();

if(PAGINAS_PUBLICAS.includes(archivo)){
return {autorizado:true,publico:true};
}

if(sessionStorage.getItem("alferza_login")!=="true"){
sessionStorage.setItem("alferza_return_url",ruta+location.search+location.hash);
location.replace(CONFIG.basePath+"login.html");
return {autorizado:false};
}

if(typeof msal==="undefined"){throw new Error("No se cargó la librería MSAL.");}

instanciaMSAL=new msal.PublicClientApplication({
auth:{clientId:CONFIG.clientId,authority:"https://login.microsoftonline.com/"+CONFIG.tenantId,redirectUri:CONFIG.redirectUri},
cache:{cacheLocation:"sessionStorage",storeAuthStateInCookie:false}
});

await instanciaMSAL.initialize();

const respuestaRedirect=await instanciaMSAL.handleRedirectPromise();

if(respuestaRedirect&&respuestaRedirect.account){
instanciaMSAL.setActiveAccount(respuestaRedirect.account);
}

const cuentas=instanciaMSAL.getAllAccounts();

if(!cuentas.length){
location.replace(CONFIG.basePath+"login.html");
return {autorizado:false};
}

const cuenta=instanciaMSAL.getActiveAccount()||cuentas[0];
instanciaMSAL.setActiveAccount(cuenta);

const respuestaToken=await instanciaMSAL.acquireTokenSilent({account:cuenta,scopes:["User.Read","Sites.Read.All"]});
tokenGraph=respuestaToken.accessToken;

const usuario=await graph("/me?$select=mail,userPrincipalName");

const correo=String(usuario.mail||usuario.userPrincipalName||cuenta.username||"").trim().toLowerCase();

if(!correo){throw new Error("No se pudo determinar el correo del usuario.");}

let permisos=leerPermisosCache(correo);

if(!permisos){
if(CONFIG.adminEmails.includes(correo)){
permisos={correo:correo,grupo:"TI",modulos:[...MODULOS]};
}else{
permisos=await obtenerPermisos(correo);
}
guardarPermisosCache(correo,permisos);
}

const moduloActual=obtenerModuloActual();

if(moduloActual==="Permisos"&&permisos.grupo!=="TI"){
throw new Error("La administración de permisos está reservada al grupo TI.");
}

if(moduloActual&&moduloActual!=="Permisos"&&!permisos.modulos.includes(moduloActual)){
throw new Error("Tu cuenta no tiene permiso para acceder al módulo "+moduloActual+".");
}

window.alferzaPermisos=permisos;
window.alferzaMSAL=instanciaMSAL;

if(document.readyState==="loading"){
await new Promise(function(resolve){document.addEventListener("DOMContentLoaded",resolve,{once:true});});
}

filtrarMenu(permisos);

window.dispatchEvent(new CustomEvent("alferza:access-granted",{detail:permisos}));

return {autorizado:true,permisos:permisos};
}

window.alferzaAccessReady=validarAcceso().catch(function(error){
console.error("ALFERZA: no se pudo validar el acceso.",error);

mostrarDenegacion("No se pudo verificar tu autorización. Comprueba tu sesión y los permisos de Microsoft 365, o contacta con el equipo de TI.",false);

return {autorizado:false,error:error.message};
});

})();