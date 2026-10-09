const msalConfig={auth:{clientId:"5d98417c-74a7-4fab-8f2c-41ac127be696",authority:"https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",redirectUri:"https://AlferzaTI.github.io/live-office/blank.html"},cache:{cacheLocation:"sessionStorage",storeAuthStateInCookie:false}};

const scopes=["User.Read","Presence.Read.All","Sites.Read.All","AuditLog.Read.All","DeviceManagementManagedDevices.Read.All"];

const msalInstance=new msal.PublicClientApplication(msalConfig);

const overlay=document.getElementById("loadingOverlay");
const loadingMessage=document.getElementById("loadingMessage");

function cambiarMensaje(mensaje){
if(loadingMessage){loadingMessage.textContent=mensaje;}
}

function mostrarCarga(mensaje){
cambiarMensaje(mensaje);
if(overlay){overlay.classList.remove("oculto");}
}

function ocultarCarga(){
if(overlay){overlay.classList.add("oculto");}
}

function esperar(ms){
return new Promise(resolve=>{setTimeout(resolve,ms);});
}

async function obtenerToken(){
const cuentas=msalInstance.getAllAccounts();
const cuenta=cuentas[0];

if(!cuenta){
throw new Error("No existe una sesión de Microsoft.");
}

msalInstance.setActiveAccount(cuenta);

try{
cambiarMensaje("Verificando permisos...");

const response=await msalInstance.acquireTokenSilent({scopes:scopes,account:cuenta});

return response.accessToken;
}catch(error){
console.error("No se pudo obtener el token silenciosamente:",error);
throw error;
}
}

async function obtenerUsuarios(TOKEN){
const respuesta=await fetch("https://graph.microsoft.com/v1.0/users?$top=999",{
method:"GET",
headers:{Authorization:`Bearer ${TOKEN}`}
});

if(!respuesta.ok){
throw new Error(`Microsoft Graph Users: HTTP ${respuesta.status}`);
}

return await respuesta.json();
}

async function obtenerPresencia(TOKEN,idsUsuarios){
let presenciaPorId={};

for(let i=0;i<idsUsuarios.length;i+=650){
const bloque=idsUsuarios.slice(i,i+650);

const respuesta=await fetch("https://graph.microsoft.com/v1.0/communications/getPresencesByUserId",{
method:"POST",
headers:{
Authorization:`Bearer ${TOKEN}`,
"Content-Type":"application/json"
},
body:JSON.stringify({ids:bloque})
});

if(!respuesta.ok){
throw new Error(`Microsoft Graph Presence: HTTP ${respuesta.status}`);
}

const data=await respuesta.json();

(data.value||[]).forEach(presencia=>{
presenciaPorId[presencia.id]=presencia;
});
}

return presenciaPorId;
}

function renderizarUsuarios(usuarios,presenciaPorId){
const contenedor=document.getElementById("officeGrid");

let disponibles=0;
let ocupados=0;
let ausentes=0;
let offline=0;

let html="";

usuarios.forEach(usuario=>{
const presencia=presenciaPorId[usuario.id]||{availability:"Offline"};
const estado=presencia.availability||"Offline";

let clase="offline";

switch(estado){
case "Available":
clase="disponible";
disponibles++;
break;

case "Busy":
case "InAMeeting":
case "OnACall":
clase="ocupado";
ocupados++;
break;

case "Away":
case "BeRightBack":
clase="ausente";
ausentes++;
break;

default:
clase="offline";
offline++;
break;
}

html+=`
<div class="card" data-estado="${estado}">
<h3>${usuario.displayName}</h3>
<p>${usuario.mail}</p>
<div class="status ${clase}">${estado}</div>
</div>
`;
});

contenedor.innerHTML=html;

document.getElementById("disp").innerText=disponibles;
document.getElementById("busy").innerText=ocupados;
document.getElementById("away").innerText=ausentes;
document.getElementById("offline").innerText=offline;
}

async function consultarInformacion(){
const TOKEN=await obtenerToken();

cambiarMensaje("Consultando personal...");

const data=await obtenerUsuarios(TOKEN);

const usuarios=data.value.filter(usuario=>usuario.mail&&usuario.mail.toLowerCase().endsWith("@alferza.pe"));

cambiarMensaje("Consultando estados...");

const idsUsuarios=usuarios.map(usuario=>usuario.id);

const presenciaPorId=await obtenerPresencia(TOKEN,idsUsuarios);

renderizarUsuarios(usuarios,presenciaPorId);
}

async function cargarUsuarios(mostrarPantalla=true){
if(mostrarPantalla){
mostrarCarga("Cargando información...");
}

try{
await consultarInformacion();

if(mostrarPantalla){
cambiarMensaje("Información actualizada");
await esperar(250);
ocultarCarga();
}

console.log("ALFERZA LIVE OFFICE actualizado correctamente.");
}catch(error){
console.error("Error cargando ALFERZA LIVE OFFICE:",error);

if(mostrarPantalla){
cambiarMensaje("No se pudo conectar con Microsoft");
await esperar(1200);
ocultarCarga();
}
}
}

cargarUsuarios(true);

setInterval(()=>{
cargarUsuarios(false);
},300000);

document.getElementById("buscador").addEventListener("keyup",function(){
const texto=this.value.toLowerCase();

document.querySelectorAll(".card").forEach(card=>{
const contenido=card.innerText.toLowerCase();
card.style.display=contenido.includes(texto)?"":"none";
});
});

let filtroActivo=null;

function filtrarEstado(tipo){
if(filtroActivo===tipo){
mostrarTodos();
filtroActivo=null;
return;
}

filtroActivo=tipo;

document.querySelectorAll(".card").forEach(card=>{
const estado=card.dataset.estado;
let mostrar=false;

switch(tipo){
case "Available":
mostrar=estado==="Available";
break;

case "Busy":
mostrar=estado==="Busy"||estado==="InAMeeting"||estado==="OnACall";
break;

case "Away":
mostrar=estado==="Away"||estado==="BeRightBack";
break;

case "Offline":
mostrar=estado==="Offline";
break;

default:
mostrar=true;
break;
}

card.style.display=mostrar?"":"none";
});
}

function mostrarTodos(){
document.querySelectorAll(".card").forEach(card=>{
card.style.display="";
});
}