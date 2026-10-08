/* global Office, msal, CONFIG */
"use strict";

const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Mail.Read", "User.Read"];

let pca = null;            // instancia MSAL
let usandoNAA = false;     // true si el host soporta Nested App Authentication
let cargaEnCurso = 0;      // contador para descartar cargas obsoletas
let adjuntosActuales = []; // [{mensaje, adjunto}] de la última carga

const $ = (id) => document.getElementById(id);

// ---------------------------------------------------------------- arranque

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Outlook) {
    mostrarError("Este complemento solo funciona dentro de Outlook.");
    return;
  }

  $("btnRecargar").addEventListener("click", () => cargar());
  $("btnLogin").addEventListener("click", () => iniciarSesion());
  $("btnTodos").addEventListener("click", () => descargarTodos());
  $("chkInline").addEventListener("change", () => pintar());

  try {
    await iniciarAuth();
  } catch (e) {
    mostrarError("No se pudo preparar la autenticación: " + mensajeDe(e));
    return;
  }

  // Al cambiar de correo con el panel anclado, recargar.
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => cargar());

  cargar();
});

// ---------------------------------------------------------------- autenticación

async function iniciarAuth() {
  if (!CONFIG || !CONFIG.clientId || CONFIG.clientId.startsWith("__")) {
    throw new Error("Falta configurar clientId/tenantId. Ejecuta configurar.py.");
  }

  const cfg = {
    auth: {
      clientId: CONFIG.clientId,
      authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
    },
  };

  if (Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
    usandoNAA = true;
    pca = await msal.createNestablePublicClientApplication(cfg);
  } else {
    // Hosts antiguos: flujo estándar con ventana emergente.
    usandoNAA = false;
    cfg.auth.redirectUri = CONFIG.redirectUri;
    cfg.cache = { cacheLocation: "localStorage" };
    pca = await msal.createStandardPublicClientApplication(cfg);
  }
}

async function obtenerToken(interactivo) {
  const req = { scopes: SCOPES };
  const cuentas = pca.getAllAccounts();
  if (cuentas.length) req.account = cuentas[0];

  try {
    const r = await pca.acquireTokenSilent(req);
    return r.accessToken;
  } catch (e) {
    if (!interactivo) throw e;
    const r = await pca.acquireTokenPopup(req);
    return r.accessToken;
  }
}

async function iniciarSesion() {
  $("login").classList.add("oculto");
  estado("Solicitando permiso…");
  try {
    await obtenerToken(true);
    cargar();
  } catch (e) {
    $("login").classList.remove("oculto");
    mostrarError("No se pudo iniciar sesión: " + mensajeDe(e));
  }
}

// ---------------------------------------------------------------- Graph

async function graph(ruta, token, opciones) {
  const url = ruta.startsWith("http") ? ruta : GRAPH + ruta;
  const resp = await fetch(url, {
    ...(opciones || {}),
    headers: { Authorization: "Bearer " + token, ...((opciones && opciones.headers) || {}) },
  });
  if (!resp.ok) {
    let detalle = resp.status + " " + resp.statusText;
    try {
      const j = await resp.json();
      if (j.error) detalle = j.error.code + ": " + j.error.message;
    } catch (_) { /* sin cuerpo JSON */ }
    const err = new Error(detalle);
    err.status = resp.status;
    throw err;
  }
  return resp;
}

async function graphJson(ruta, token) {
  return (await graph(ruta, token)).json();
}

// ---------------------------------------------------------------- carga del hilo

async function cargar() {
  const item = Office.context.mailbox.item;
  if (!item || !item.itemId) {
    estado("Selecciona un correo.");
    $("lista").innerHTML = "";
    return;
  }

  const miCarga = ++cargaEnCurso;
  adjuntosActuales = [];
  $("btnTodos").disabled = true;
  $("lista").innerHTML = "";
  $("asunto").textContent = item.subject || "";
  $("resumen").textContent = "";
  estado("Leyendo conversación…");

  let token;
  try {
    token = await obtenerToken(false);
  } catch (_) {
    estado("");
    $("login").classList.remove("oculto");
    return;
  }
  $("login").classList.add("oculto");

  try {
    // 1) Id REST del correo abierto -> conversationId según Graph
    const restId = Office.context.mailbox.convertToRestId(
      item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
    const actual = await graphJson(
      "/me/messages/" + encodeURIComponent(restId) + "?$select=id,conversationId,subject",
      token
    );
    if (miCarga !== cargaEnCurso) return;

    // 2) Todos los mensajes de la conversación (todas las carpetas, incl. Enviados)
    const filtro = encodeURIComponent("conversationId eq '" + actual.conversationId + "'");
    let ruta =
      "/me/messages?$filter=" + filtro +
      "&$select=id,subject,from,sender,receivedDateTime,sentDateTime,hasAttachments,isDraft,webLink" +
      "&$top=100";
    const mensajes = [];
    while (ruta) {
      const pagina = await graphJson(ruta, token);
      mensajes.push(...pagina.value);
      ruta = pagina["@odata.nextLink"] || null;
    }
    if (miCarga !== cargaEnCurso) return;

    mensajes.sort((a, b) => new Date(fechaDe(b)) - new Date(fechaDe(a)));

    // 3) Adjuntos de cada mensaje que los tenga
    estado("Buscando adjuntos en " + mensajes.length + " mensaje(s)…");
    const conAdjuntos = mensajes.filter((m) => m.hasAttachments && !m.isDraft);
    const listas = await Promise.all(
      conAdjuntos.map((m) =>
        graphJson(
          "/me/messages/" + encodeURIComponent(m.id) +
            "/attachments?$select=id,name,contentType,size,isInline",
          token
        ).then((r) => r.value).catch(() => [])
      )
    );
    if (miCarga !== cargaEnCurso) return;

    conAdjuntos.forEach((m, i) => {
      listas[i].forEach((a) => adjuntosActuales.push({ mensaje: m, adjunto: a, actual: m.id === actual.id }));
    });

    estado("");
    pintar(mensajes.length);
  } catch (e) {
    if (miCarga !== cargaEnCurso) return;
    if (e.status === 401 || e.status === 403) {
      $("login").classList.remove("oculto");
      mostrarError("Sin permiso para leer el buzón: " + mensajeDe(e));
    } else {
      mostrarError("Error al leer la conversación: " + mensajeDe(e));
    }
  }
}

// ---------------------------------------------------------------- render

function pintar(nMensajes) {
  const verInline = $("chkInline").checked;
  const visibles = adjuntosActuales.filter((x) => verInline || !x.adjunto.isInline);

  const lista = $("lista");
  lista.innerHTML = "";

  if (!visibles.length) {
    lista.innerHTML = '<div class="vacio">No hay archivos adjuntos en esta conversación.</div>';
    $("btnTodos").disabled = true;
    $("resumen").textContent = nMensajes != null ? nMensajes + " mensaje(s) en el hilo" : "";
    return;
  }

  // Agrupar por mensaje conservando el orden (más reciente primero)
  const grupos = new Map();
  for (const x of visibles) {
    if (!grupos.has(x.mensaje.id)) grupos.set(x.mensaje.id, { mensaje: x.mensaje, actual: x.actual, adjuntos: [] });
    grupos.get(x.mensaje.id).adjuntos.push(x.adjunto);
  }

  for (const g of grupos.values()) {
    const bloque = document.createElement("section");
    bloque.className = "mensaje";

    const cab = document.createElement("div");
    cab.className = "mensaje-cab";
    const quien = remitenteDe(g.mensaje);
    cab.innerHTML =
      '<div><span class="quien"></span>' +
      (g.actual ? '<span class="actual">· este correo</span>' : "") +
      '<div class="cuando"></div></div>';
    cab.querySelector(".quien").textContent = quien;
    cab.querySelector(".cuando").textContent = formatearFecha(fechaDe(g.mensaje));

    const btnAbrir = document.createElement("button");
    btnAbrir.className = "btn secundario mini";
    btnAbrir.textContent = "Abrir correo";
    btnAbrir.addEventListener("click", () => abrirMensaje(g.mensaje));
    cab.appendChild(btnAbrir);
    bloque.appendChild(cab);

    for (const a of g.adjuntos) bloque.appendChild(filaAdjunto(g.mensaje, a));
    lista.appendChild(bloque);
  }

  const descargables = visibles.filter((x) => esArchivo(x.adjunto)).length;
  $("btnTodos").disabled = descargables === 0;
  $("resumen").textContent =
    visibles.length + " adjunto(s) en " + grupos.size + " mensaje(s)" +
    (nMensajes != null ? " · " + nMensajes + " mensaje(s) en el hilo" : "");
}

function filaAdjunto(m, a) {
  const fila = document.createElement("div");
  fila.className = "adjunto";

  const nombre = document.createElement("span");
  nombre.className = "nombre";
  nombre.title = a.name || "";
  nombre.textContent = a.name || "(sin nombre)";
  fila.appendChild(nombre);

  const tam = document.createElement("span");
  tam.className = "tam";
  tam.textContent = formatearTamano(a.size);
  fila.appendChild(tam);

  if (esArchivo(a)) {
    const btn = document.createElement("button");
    btn.className = "btn mini";
    btn.textContent = "Descargar";
    btn.addEventListener("click", () => descargar(m, a, btn));
    fila.appendChild(btn);
  } else if (a["@odata.type"] === "#microsoft.graph.referenceAttachment") {
    const tipo = document.createElement("span");
    tipo.className = "tipo";
    tipo.textContent = "enlace OneDrive/SharePoint";
    fila.appendChild(tipo);
    const btn = document.createElement("button");
    btn.className = "btn secundario mini";
    btn.textContent = "Abrir correo";
    btn.addEventListener("click", () => abrirMensaje(m));
    fila.appendChild(btn);
  } else {
    const tipo = document.createElement("span");
    tipo.className = "tipo";
    tipo.textContent = "correo adjunto";
    fila.appendChild(tipo);
    const btn = document.createElement("button");
    btn.className = "btn secundario mini";
    btn.textContent = "Abrir correo";
    btn.addEventListener("click", () => abrirMensaje(m));
    fila.appendChild(btn);
  }
  return fila;
}

// ---------------------------------------------------------------- acciones

async function descargar(m, a, btn) {
  const textoOriginal = btn ? btn.textContent : "";
  if (btn) { btn.disabled = true; btn.textContent = "…"; }
  try {
    const token = await obtenerToken(false);
    const resp = await graph(
      "/me/messages/" + encodeURIComponent(m.id) + "/attachments/" + encodeURIComponent(a.id) + "/$value",
      token
    );
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const enlace = document.createElement("a");
    enlace.href = url;
    enlace.download = a.name || "adjunto";
    document.body.appendChild(enlace);
    enlace.click();
    enlace.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  } catch (e) {
    mostrarError("No se pudo descargar «" + a.name + "»: " + mensajeDe(e));
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = textoOriginal; }
  }
}

async function descargarTodos() {
  const verInline = $("chkInline").checked;
  const lista = adjuntosActuales.filter((x) => esArchivo(x.adjunto) && (verInline || !x.adjunto.isInline));
  $("btnTodos").disabled = true;
  let n = 0;
  for (const x of lista) {
    estado("Descargando " + (++n) + " de " + lista.length + "…");
    await descargar(x.mensaje, x.adjunto, null);
    await new Promise((r) => setTimeout(r, 400)); // evitar que el navegador agrupe descargas
  }
  estado("");
  $("btnTodos").disabled = false;
}

function abrirMensaje(m) {
  try {
    const ewsId = Office.context.mailbox.convertToEwsId(m.id, Office.MailboxEnums.RestVersion.v2_0);
    Office.context.mailbox.displayMessageForm(ewsId);
  } catch (_) {
    if (m.webLink) window.open(m.webLink, "_blank");
  }
}

// ---------------------------------------------------------------- utilidades

function esArchivo(a) {
  return a["@odata.type"] === "#microsoft.graph.fileAttachment";
}

function fechaDe(m) {
  return m.receivedDateTime || m.sentDateTime || "";
}

function remitenteDe(m) {
  const e = (m.from && m.from.emailAddress) || (m.sender && m.sender.emailAddress) || {};
  return e.name || e.address || "(desconocido)";
}

function formatearFecha(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  return d.toLocaleDateString("es-ES", { day: "2-digit", month: "2-digit", year: "numeric" }) +
    " " + d.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" });
}

function formatearTamano(bytes) {
  if (bytes == null) return "";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(0) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function mensajeDe(e) {
  return (e && (e.errorMessage || e.message)) || String(e);
}

function estado(texto) {
  const el = $("estado");
  el.classList.remove("error");
  el.textContent = texto || "";
}

function mostrarError(texto) {
  const el = $("estado");
  el.classList.add("error");
  el.textContent = texto;
}
