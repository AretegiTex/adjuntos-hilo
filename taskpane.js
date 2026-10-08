/* global Office, msal, CONFIG */
"use strict";

const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Mail.Read", "User.Read"];
const MAX_ADJUNTAR = 25 * 1024 * 1024;   // límite de Outlook para adjuntar desde base64 (~25 MB)
const MAX_PREVIEW = 15 * 1024 * 1024;    // tamaño máximo que previsualizamos
const MAX_CACHE = 80 * 1024 * 1024;      // memoria total de archivos preparados

let pca = null;            // instancia MSAL
let cargaEnCurso = 0;      // contador para descartar cargas obsoletas
let adjuntosActuales = []; // [{mensaje, adjunto, actual}] de la última carga
let nMensajesHilo = 0;
let modo = "read";         // "read" | "compose"
let simulado = false;      // vista previa de diseño fuera de Outlook
let previewAbierta = null; // id del adjunto con vista previa desplegada
const cache = new Map();   // id adjunto -> {blob, url, size, t}

const $ = (id) => document.getElementById(id);

const ICONO = {
  descargar:
    '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="M8 2.5v8"/><path d="m5 7.5 3 3 3-3"/><path d="M2.5 11.5v1.5a.5.5 0 0 0 .5.5h10a.5.5 0 0 0 .5-.5v-1.5"/></svg>',
  adjuntar:
    '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="m10.3 5.2-4.4 4.4a1.3 1.3 0 0 0 1.9 1.9l4.8-4.8a2.6 2.6 0 0 0-3.7-3.7L4 7.9a3.9 3.9 0 0 0 5.5 5.5l4-4"/></svg>',
  ojo:
    '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="M1.5 8s2.5-4.5 6.5-4.5S14.5 8 14.5 8 12 12.5 8 12.5 1.5 8 1.5 8z"/><circle cx="8" cy="8" r="2"/></svg>',
  cerrar:
    '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" aria-hidden="true">' +
    '<path d="m4 4 8 8M12 4l-8 8"/></svg>',
  ok:
    '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="m3 8.5 3 3 7-7"/></svg>',
  clip:
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
    '<path d="M15.5 7.5 8.6 14.4a2 2 0 0 0 2.8 2.8l7.4-7.4a4 4 0 0 0-5.7-5.7L5.4 11.8a6 6 0 0 0 8.5 8.5l6.1-6.1"/></svg>',
};

// ---------------------------------------------------------------- arranque

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(async (info) => {
    if (simulado) return;
    if (info.host !== Office.HostType.Outlook) {
      mostrarError("Este panel solo funciona dentro de Outlook.");
      return;
    }

    conectarControles();
    detectarModo();

    try {
      await iniciarAuth();
    } catch (e) {
      mostrarError("No se pudo preparar la autenticación. " + mensajeDe(e));
      return;
    }

    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { detectarModo(); cargar(); });
    cargar();
  });
}

function detectarModo() {
  const item = Office.context.mailbox.item;
  modo = item && typeof item.addFileAttachmentFromBase64Async === "function" ? "compose" : "read";
  document.body.classList.toggle("compose", modo === "compose");
}

function conectarControles() {
  $("btnRecargar").addEventListener("click", () => cargar());
  $("btnLogin").addEventListener("click", () => iniciarSesion());
  $("btnTodos").addEventListener("click", () => accionTodos());
  $("chkInline").addEventListener("change", () => pintar());
  $("chkVersiones").addEventListener("change", () => pintar());
  $("filtro").addEventListener("input", () => pintar());
}

// ---------------------------------------------------------------- autenticación

async function iniciarAuth() {
  if (!CONFIG || !CONFIG.clientId || CONFIG.clientId.startsWith("__")) {
    throw new Error("Falta configurar clientId y tenantId: ejecuta configurar.py y sube config.js.");
  }
  const cfg = {
    auth: {
      clientId: CONFIG.clientId,
      authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
    },
  };
  if (Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
    pca = await msal.createNestablePublicClientApplication(cfg);
  } else {
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
    return (await pca.acquireTokenSilent(req)).accessToken;
  } catch (e) {
    if (!interactivo) throw e;
    return (await pca.acquireTokenPopup(req)).accessToken;
  }
}

async function iniciarSesion() {
  $("login").classList.add("oculto");
  limpiarError();
  try {
    await obtenerToken(true);
    cargar();
  } catch (e) {
    $("login").classList.remove("oculto");
    mostrarError("No se pudo iniciar sesión. " + mensajeDe(e));
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

const graphJson = async (ruta, token) => (await graph(ruta, token)).json();

async function mensajesDeConversacion(convId, token) {
  const filtro = encodeURIComponent("conversationId eq '" + convId + "'");
  let ruta =
    "/me/messages?$filter=" + filtro +
    "&$select=id,subject,from,sender,receivedDateTime,sentDateTime,hasAttachments,isDraft,webLink&$top=100";
  const mensajes = [];
  while (ruta) {
    const pagina = await graphJson(ruta, token);
    mensajes.push(...pagina.value);
    ruta = pagina["@odata.nextLink"] || null;
  }
  return mensajes;
}

// ---------------------------------------------------------------- identificar la conversación

// Lectura: el correo abierto tiene itemId -> Graph nos da su conversationId.
async function conversacionEnLectura(item, token) {
  const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  const actual = await graphJson("/me/messages/" + encodeURIComponent(restId) + "?$select=id,conversationId", token);
  return { convId: actual.conversationId, idActual: actual.id };
}

// Redacción: una respuesta ya trae conversationId (formato EWS); lo pasamos al
// formato de Graph y, si no cuadra, guardamos el borrador y preguntamos por él.
async function conversacionEnRedaccion(item, token) {
  const bruto = item.conversationId;
  if (bruto) {
    const candidatos = [bruto.replace(/\//g, "-").replace(/\+/g, "_"), bruto];
    for (const c of candidatos) {
      try {
        const lista = await mensajesDeConversacion(c, token);
        if (lista.length) return { convId: c, idActual: null, mensajes: lista };
      } catch (_) { /* probamos el siguiente formato */ }
    }
  }
  const itemId = await new Promise((res, rej) =>
    item.saveAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(new Error(r.error && r.error.message))))
  );
  const restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
  const borrador = await graphJson("/me/messages/" + encodeURIComponent(restId) + "?$select=id,conversationId", token);
  return { convId: borrador.conversationId, idActual: borrador.id };
}

// ---------------------------------------------------------------- carga del hilo

async function cargar() {
  const item = Office.context.mailbox.item;
  if (!item) {
    $("lista").innerHTML = "";
    $("asunto").textContent = "";
    $("resumen").textContent = "Selecciona un correo para ver los archivos de su conversación.";
    return;
  }

  const miCarga = ++cargaEnCurso;
  adjuntosActuales = [];
  nMensajesHilo = 0;
  previewAbierta = null;
  limpiarError();
  $("btnTodos").disabled = true;
  $("btnRecargar").classList.add("spin");
  $("resumen").textContent = "Leyendo la conversación";
  $("pie").textContent = "";
  pintarEsqueleto();

  if (modo === "compose") {
    item.subject.getAsync((r) => { if (miCarga === cargaEnCurso) $("asunto").textContent = r.value || ""; });
  } else {
    $("asunto").textContent = item.subject || "";
  }

  let token;
  try {
    token = await obtenerToken(false);
  } catch (_) {
    $("btnRecargar").classList.remove("spin");
    $("lista").innerHTML = "";
    $("resumen").textContent = "";
    $("login").classList.remove("oculto");
    return;
  }
  $("login").classList.add("oculto");

  try {
    const conv = modo === "compose"
      ? await conversacionEnRedaccion(item, token)
      : await conversacionEnLectura(item, token);
    if (miCarga !== cargaEnCurso) return;

    const mensajes = conv.mensajes || await mensajesDeConversacion(conv.convId, token);
    if (miCarga !== cargaEnCurso) return;

    mensajes.sort((a, b) => new Date(fechaDe(b)) - new Date(fechaDe(a)));
    nMensajesHilo = mensajes.length;

    const conAdjuntos = mensajes.filter((m) => m.hasAttachments && !m.isDraft);
    const listas = await Promise.all(
      conAdjuntos.map((m) =>
        graphJson(
          "/me/messages/" + encodeURIComponent(m.id) + "/attachments?$select=id,name,contentType,size,isInline",
          token
        ).then((r) => r.value).catch(() => [])
      )
    );
    if (miCarga !== cargaEnCurso) return;

    conAdjuntos.forEach((m, i) => {
      listas[i].forEach((a) => adjuntosActuales.push({ mensaje: m, adjunto: a, actual: m.id === conv.idActual }));
    });

    pintar();
  } catch (e) {
    if (miCarga !== cargaEnCurso) return;
    $("lista").innerHTML = "";
    $("resumen").textContent = "";
    if (e.status === 401 || e.status === 403) {
      $("login").classList.remove("oculto");
      mostrarError("No hay permiso para leer el buzón. " + mensajeDe(e));
    } else {
      mostrarError("No se pudo leer la conversación. " + mensajeDe(e));
    }
  } finally {
    if (miCarga === cargaEnCurso) $("btnRecargar").classList.remove("spin");
  }
}

// ---------------------------------------------------------------- selección (filtro, inline, versiones)

function seleccionar() {
  const verInline = $("chkInline").checked;
  const soloUltimas = $("chkVersiones").checked;
  const texto = normalizar($("filtro").value.trim());

  let lista = adjuntosActuales.filter((x) => verInline || !x.adjunto.isInline);
  if (texto) lista = lista.filter((x) => normalizar(x.adjunto.name || "").includes(texto));

  const porNombre = new Map();
  for (const x of lista) {
    const clave = normalizar(x.adjunto.name || "");
    if (!porNombre.has(clave)) porNombre.set(clave, []);
    porNombre.get(clave).push(x);
  }
  let ocultas = 0;
  for (const grupo of porNombre.values()) {
    grupo.sort((a, b) => new Date(fechaDe(b.mensaje)) - new Date(fechaDe(a.mensaje)));
    grupo.forEach((x, i) => {
      x.anteriores = i === 0 ? grupo.length - 1 : 0;
      x.esAnterior = i > 0;
    });
    if (soloUltimas) ocultas += grupo.length - 1;
  }
  if (soloUltimas) lista = lista.filter((x) => !x.esAnterior);

  return { lista, ocultas, filtrando: Boolean(texto) };
}

// ---------------------------------------------------------------- render

function pintar() {
  const { lista, ocultas, filtrando } = seleccionar();
  const cont = $("lista");
  cont.innerHTML = "";
  cont.classList.remove("skeleton");

  const total = adjuntosActuales.filter((x) => $("chkInline").checked || !x.adjunto.isInline).length;

  const partes = [];
  if (adjuntosActuales.length === 0) {
    partes.push(nMensajesHilo ? "Sin archivos en " + plural(nMensajesHilo, "mensaje") : "");
  } else if (filtrando) {
    partes.push(lista.length + " de " + plural(total, "archivo"));
  } else {
    const nMsgs = new Set(lista.map((x) => x.mensaje.id)).size;
    partes.push(plural(lista.length, "archivo") + " en " + plural(nMsgs, "mensaje"));
  }
  if (ocultas) partes.push(plural(ocultas, "versión anterior", "versiones anteriores") + " ocultas");
  $("resumen").textContent = partes.filter(Boolean).join(", ");

  if (!lista.length) {
    cont.appendChild(vacio(filtrando));
    actualizarBotonTodos(0);
    return;
  }

  const grupos = new Map();
  for (const x of lista) {
    if (!grupos.has(x.mensaje.id)) grupos.set(x.mensaje.id, { mensaje: x.mensaje, actual: x.actual, items: [] });
    grupos.get(x.mensaje.id).items.push(x);
  }

  for (const g of grupos.values()) {
    const sec = document.createElement("section");
    sec.className = "msg" + (g.actual ? " current" : "");

    const rail = document.createElement("div");
    rail.className = "rail";
    rail.innerHTML = '<span class="node"></span>';
    sec.appendChild(rail);

    const body = document.createElement("div");
    body.className = "msg-body";

    const head = document.createElement("div");
    head.className = "msg-head";
    const who = document.createElement("span");
    who.className = "who";
    who.textContent = remitenteDe(g.mensaje);
    who.title = direccionDe(g.mensaje);
    const when = document.createElement("span");
    when.className = "when";
    when.textContent = formatearFecha(fechaDe(g.mensaje));
    const abrir = document.createElement("button");
    abrir.className = "open-link";
    abrir.textContent = "Abrir correo";
    abrir.addEventListener("click", () => abrirMensaje(g.mensaje));
    head.append(who, when, abrir);
    body.appendChild(head);

    const ul = document.createElement("ul");
    ul.className = "files";
    for (const x of g.items) ul.appendChild(filaAdjunto(x));
    body.appendChild(ul);

    sec.appendChild(body);
    cont.appendChild(sec);
  }

  actualizarBotonTodos(lista.filter((x) => esArchivo(x.adjunto)).length);

  const hayArchivos = lista.some((x) => esArchivo(x.adjunto));
  $("pie").textContent = hayArchivos && modo === "compose"
    ? "Pulsa el clip de un archivo para adjuntarlo a este correo."
    : "";
}

function actualizarBotonTodos(n) {
  const verbo = modo === "compose" ? "Adjuntar todo" : "Descargar todo";
  $("btnTodos").disabled = n === 0;
  $("btnTodosTexto").textContent = n > 1 ? verbo + " (" + n + ")" : verbo;
}

function filaAdjunto(x) {
  const a = x.adjunto;
  const m = x.mensaje;
  const li = document.createElement("li");
  li.className = "file" + (x.esAnterior ? " previous" : "");
  li.dataset.id = a.id;

  const fila = document.createElement("div");
  fila.className = "file-row";

  const tipo = tipoDe(a);
  const tile = document.createElement("span");
  tile.className = "tile tile-" + tipo.clase;
  tile.textContent = tipo.etiqueta;
  tile.setAttribute("aria-hidden", "true");

  const texto = document.createElement("div");
  texto.className = "file-text";
  const nombre = document.createElement("span");
  nombre.className = "file-name";
  nombre.textContent = a.name || "(sin nombre)";
  nombre.title = a.name || "";
  const meta = document.createElement("span");
  meta.className = "file-meta";
  meta.appendChild(etiqueta("size", formatearTamano(a.size)));
  if (a["@odata.type"] === "#microsoft.graph.referenceAttachment") meta.appendChild(etiqueta("tag", "Enlace a OneDrive o SharePoint"));
  else if (a["@odata.type"] === "#microsoft.graph.itemAttachment") meta.appendChild(etiqueta("tag", "Correo adjunto"));
  if (x.anteriores) meta.appendChild(etiqueta("versions", plural(x.anteriores, "versión anterior", "versiones anteriores")));
  if (x.esAnterior) meta.appendChild(etiqueta("tag", "Versión anterior"));
  texto.append(nombre, meta);

  const acciones = document.createElement("div");
  acciones.className = "file-actions";
  if (esArchivo(a)) {
    if (previsualizable(a)) {
      const btnVer = botonIcono(ICONO.ojo, "Vista previa");
      btnVer.addEventListener("click", () => togglePreview(x, li, btnVer));
      acciones.appendChild(btnVer);
    }
    if (modo === "compose") {
      const btnAdj = botonIcono(ICONO.adjuntar, "Adjuntar a este correo");
      btnAdj.classList.add("attach");
      btnAdj.addEventListener("click", () => adjuntarAlCorreo(x, btnAdj));
      acciones.appendChild(btnAdj);
    }
    const btnDesc = botonIcono(ICONO.descargar, "Descargar");
    btnDesc.addEventListener("click", () => descargar(x, btnDesc));
    acciones.appendChild(btnDesc);

  }

  fila.append(tile, texto, acciones);
  li.appendChild(fila);
  return li;
}

function botonIcono(svg, titulo) {
  const b = document.createElement("button");
  b.className = "icon-btn";
  b.title = titulo;
  b.setAttribute("aria-label", titulo);
  b.innerHTML = svg;
  return b;
}

function etiqueta(clase, texto) {
  const s = document.createElement("span");
  s.className = clase;
  s.textContent = texto;
  return s;
}

function vacio(filtrando) {
  const div = document.createElement("div");
  div.className = "empty";
  div.innerHTML = ICONO.clip;
  const p = document.createElement("p");
  p.textContent = filtrando ? "Ningún archivo coincide con la búsqueda." : "Esta conversación no tiene archivos adjuntos.";
  div.appendChild(p);
  if (!filtrando && adjuntosActuales.some((x) => x.adjunto.isInline)) {
    const h = document.createElement("p");
    h.className = "hint";
    h.textContent = "Hay imágenes incrustadas; actívalas arriba si las necesitas.";
    div.appendChild(h);
  }
  return div;
}

function pintarEsqueleto() {
  const cont = $("lista");
  cont.classList.add("skeleton");
  cont.innerHTML = "";
  const hueco = (n) => "&nbsp;".repeat(n);
  [2, 1].forEach((n, i) => {
    const sec = document.createElement("section");
    sec.className = "msg" + (i === 0 ? " current" : "");
    sec.innerHTML =
      '<div class="rail"><span class="node"></span></div>' +
      '<div class="msg-body"><div class="msg-head"><span class="who">' + hueco(18) + '</span>' +
      '<span class="when">' + hueco(12) + '</span></div><ul class="files">' +
      Array.from({ length: n }, () =>
        '<li class="file"><div class="file-row"><span class="tile"></span><div class="file-text">' +
        '<span class="file-name">' + hueco(30) + '</span><span class="file-meta">' + hueco(8) + '</span></div></div></li>'
      ).join("") +
      "</ul></div>";
    cont.appendChild(sec);
  });
}

// ---------------------------------------------------------------- contenido de archivos (caché)

async function obtenerBlob(x) {
  const a = x.adjunto;
  const hit = cache.get(a.id);
  if (hit) { hit.t = Date.now(); return hit; }
  const token = await obtenerToken(false);
  const resp = await graph(
    "/me/messages/" + encodeURIComponent(x.mensaje.id) + "/attachments/" + encodeURIComponent(a.id) + "/$value",
    token
  );
  let blob = await resp.blob();
  const tipo = a.contentType || blob.type || "application/octet-stream";
  if (blob.type !== tipo) blob = new Blob([blob], { type: tipo });
  const entrada = { blob, url: URL.createObjectURL(blob), size: blob.size, t: Date.now() };
  cache.set(a.id, entrada);
  limpiarCache();
  return entrada;
}

function limpiarCache() {
  let total = 0;
  for (const e of cache.values()) total += e.size;
  if (total <= MAX_CACHE) return;
  const orden = [...cache.entries()].sort((p, q) => p[1].t - q[1].t);
  for (const [id, e] of orden) {
    if (total <= MAX_CACHE || id === previewAbierta) continue;
    URL.revokeObjectURL(e.url);
    cache.delete(id);
    total -= e.size;
  }
}

function blobABase64(blob) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result).split(",")[1] || "");
    r.onerror = () => rej(r.error);
    r.readAsDataURL(blob);
  });
}

// ---------------------------------------------------------------- descargar

async function descargar(x, btn) {
  const a = x.adjunto;
  if (btn) btn.disabled = true;
  try {
    const e = await obtenerBlob(x);
    const enlace = document.createElement("a");
    enlace.href = e.url;
    enlace.download = a.name || "adjunto";
    document.body.appendChild(enlace);
    enlace.click();
    enlace.remove();
  } catch (err) {
    mostrarError("No se pudo descargar «" + a.name + "». " + mensajeDe(err));
  } finally {
    if (btn) btn.disabled = false;
  }
}

// ---------------------------------------------------------------- adjuntar al correo en redacción

function adjuntarBase64(base64, nombre) {
  return new Promise((res, rej) => {
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(base64, nombre, { isInline: false }, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) res(r.value);
      else rej(new Error((r.error && r.error.message) || "Outlook rechazó el adjunto."));
    });
  });
}

async function adjuntarAlCorreo(x, btn) {
  const a = x.adjunto;
  if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
    mostrarError("Esta versión de Outlook no permite adjuntar desde el panel.");
    return;
  }
  if ((a.size || 0) > MAX_ADJUNTAR) {
    mostrarError("«" + a.name + "» supera el máximo que Outlook admite desde el panel (25 MB). Descárgalo y adjúntalo a mano.");
    return;
  }
  if (btn) { btn.disabled = true; btn.classList.add("working"); }
  try {
    const e = await obtenerBlob(x);
    const b64 = await blobABase64(e.blob);
    await adjuntarBase64(b64, a.name || "adjunto");
    if (btn) { btn.classList.remove("working"); btn.classList.add("done"); btn.innerHTML = ICONO.ok; btn.title = "Adjuntado"; }
  } catch (err) {
    if (btn) { btn.disabled = false; btn.classList.remove("working"); }
    mostrarError("No se pudo adjuntar «" + a.name + "». " + mensajeDe(err));
  }
}

async function accionTodos() {
  const { lista } = seleccionar();
  const archivos = lista.filter((x) => esArchivo(x.adjunto));
  const btn = $("btnTodos");
  const texto = $("btnTodosTexto");
  btn.disabled = true;
  let n = 0;
  for (const x of archivos) {
    n++;
    texto.textContent = (modo === "compose" ? "Adjuntando " : "Descargando ") + n + " de " + archivos.length;
    const fila = $("lista").querySelector('.file[data-id="' + CSS.escape(x.adjunto.id) + '"]');
    if (modo === "compose") {
      await adjuntarAlCorreo(x, fila ? fila.querySelector(".attach") : null);
    } else {
      await descargar(x, null);
      await new Promise((r) => setTimeout(r, 400));
    }
  }
  actualizarBotonTodos(archivos.length);
}

// ---------------------------------------------------------------- vista previa

function previsualizable(a) {
  if (!esArchivo(a)) return false;
  if ((a.size || 0) > MAX_PREVIEW) return false;
  const t = (a.contentType || "").toLowerCase();
  const n = (a.name || "").toLowerCase();
  return t.startsWith("image/") || t === "application/pdf" || /\.(pdf|png|jpe?g|gif|webp|bmp|svg)$/.test(n);
}

function esImagen(a) {
  const t = (a.contentType || "").toLowerCase();
  return t.startsWith("image/") || /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(a.name || "");
}

async function togglePreview(x, li, btn) {
  const a = x.adjunto;
  const existente = li.querySelector(".preview");
  if (existente) {
    existente.remove();
    li.classList.remove("open");
    previewAbierta = null;
    return;
  }
  // una sola vista previa abierta a la vez
  $("lista").querySelectorAll(".file .preview").forEach((p) => { p.closest(".file").classList.remove("open"); p.remove(); });
  previewAbierta = a.id;

  const caja = document.createElement("div");
  caja.className = "preview";
  caja.innerHTML = '<div class="preview-estado">Cargando vista previa</div>';
  li.appendChild(caja);
  li.classList.add("open");
  if (simulado) {
    caja.innerHTML = '<div class="preview-estado">Aquí se muestra el PDF o la imagen (solo dentro de Outlook).</div>' +
      '<div class="preview-bar"><button class="btn quiet">Ampliar</button><button class="btn quiet" onclick="this.closest(\'.file\').classList.remove(\'open\');this.closest(\'.preview\').remove()">Cerrar</button></div>';
    return;
  }
  btn.disabled = true;
  try {
    const e = await obtenerBlob(x);
    if (previewAbierta !== a.id) return;
    caja.innerHTML = "";
    if (esImagen(a)) {
      const img = document.createElement("img");
      img.src = e.url;
      img.alt = a.name || "";
      caja.appendChild(img);
    } else {
      const pdf = document.createElement("div");
      pdf.className = "pdf";
      pdf.innerHTML = '<div class="preview-estado">Preparando el PDF</div>';
      caja.appendChild(pdf);
      const idAbierta = a.id;
      const total = await window.renderizarPdf(e.blob, pdf, {
        ancho: pdf.clientWidth || 300,
        maxPaginas: 3,
        cancelado: () => previewAbierta !== idAbierta,
      });
      if (total > 3) {
        const mas = document.createElement("div");
        mas.className = "preview-estado";
        mas.textContent = "Mostrando 3 de " + total + " páginas. Ampliar para ver el documento completo.";
        pdf.appendChild(mas);
      }
    }
    if (previewAbierta !== a.id) return;
    const barra = document.createElement("div");
    barra.className = "preview-bar";
    if (puedeAmpliar()) {
      const ampliar = document.createElement("button");
      ampliar.className = "btn quiet";
      ampliar.textContent = "Ampliar";
      ampliar.addEventListener("click", () => ampliarPreview(x, ampliar));
      barra.appendChild(ampliar);
    }
    const cerrar = document.createElement("button");
    cerrar.className = "btn quiet";
    cerrar.textContent = "Cerrar";
    cerrar.addEventListener("click", () => togglePreview(x, li, btn));
    barra.appendChild(cerrar);
    caja.appendChild(barra);
  } catch (err) {
    caja.innerHTML = '<div class="preview-estado">No se pudo cargar la vista previa.</div>';
    mostrarError("No se pudo cargar «" + a.name + "». " + mensajeDe(err));
  } finally {
    btn.disabled = false;
  }
}

function puedeAmpliar() {
  return !simulado &&
    typeof Office !== "undefined" && Office.context && Office.context.ui &&
    Office.context.requirements.isSetSupported("DialogApi", "1.2");
}

async function ampliarPreview(x, btn) {
  const a = x.adjunto;
  btn.disabled = true;
  try {
    const e = await obtenerBlob(x);
    const b64 = await blobABase64(e.blob);
    const url = new URL("visor.html", location.href).href;
    Office.context.ui.displayDialogAsync(url, { height: 85, width: 70, displayInIframe: true }, (r) => {
      btn.disabled = false;
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        mostrarError("No se pudo abrir la ventana de vista previa. " + ((r.error && r.error.message) || ""));
        return;
      }
      const dlg = r.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "listo") {
          dlg.messageChild(JSON.stringify({ name: a.name, type: e.blob.type || a.contentType, base64: b64 }));
        }
      });
    });
  } catch (err) {
    btn.disabled = false;
    mostrarError("No se pudo ampliar «" + a.name + "». " + mensajeDe(err));
  }
}

// ---------------------------------------------------------------- abrir correo

function abrirMensaje(m) {
  if (simulado) return;
  try {
    const ewsId = Office.context.mailbox.convertToEwsId(m.id, Office.MailboxEnums.RestVersion.v2_0);
    Office.context.mailbox.displayMessageForm(ewsId);
  } catch (_) {
    if (m.webLink) window.open(m.webLink, "_blank");
  }
}

// ---------------------------------------------------------------- utilidades

const esArchivo = (a) => a["@odata.type"] === "#microsoft.graph.fileAttachment";
const fechaDe = (m) => m.receivedDateTime || m.sentDateTime || "";

function remitenteDe(m) {
  const e = (m.from && m.from.emailAddress) || (m.sender && m.sender.emailAddress) || {};
  return e.name || e.address || "Remitente desconocido";
}
function direccionDe(m) {
  const e = (m.from && m.from.emailAddress) || (m.sender && m.sender.emailAddress) || {};
  return e.address || "";
}

function tipoDe(a) {
  if (a["@odata.type"] === "#microsoft.graph.itemAttachment") return { clase: "otro", etiqueta: "EML" };
  if (a["@odata.type"] === "#microsoft.graph.referenceAttachment") return { clase: "otro", etiqueta: "URL" };
  const nombre = (a.name || "").toLowerCase();
  const ext = nombre.includes(".") ? nombre.split(".").pop() : "";
  const tabla = {
    pdf: ["pdf", "PDF"],
    xls: ["xls", "XLS"], xlsx: ["xls", "XLSX"], xlsm: ["xls", "XLSM"], csv: ["xls", "CSV"],
    doc: ["doc", "DOC"], docx: ["doc", "DOCX"], rtf: ["doc", "RTF"], odt: ["doc", "ODT"],
    ppt: ["ppt", "PPT"], pptx: ["ppt", "PPTX"],
    jpg: ["img", "JPG"], jpeg: ["img", "JPG"], png: ["img", "PNG"], gif: ["img", "GIF"], bmp: ["img", "BMP"],
    heic: ["img", "HEIC"], webp: ["img", "WEBP"], tif: ["img", "TIF"], tiff: ["img", "TIF"], svg: ["img", "SVG"],
    zip: ["zip", "ZIP"], rar: ["zip", "RAR"], "7z": ["zip", "7Z"],
    dwg: ["cad", "DWG"], dxf: ["cad", "DXF"], ifc: ["cad", "IFC"], rvt: ["cad", "RVT"],
    msg: ["otro", "MSG"], eml: ["otro", "EML"], txt: ["otro", "TXT"], xml: ["otro", "XML"],
  };
  if (tabla[ext]) return { clase: tabla[ext][0], etiqueta: tabla[ext][1] };
  return { clase: "otro", etiqueta: ext ? ext.slice(0, 4).toUpperCase() : "—" };
}

function normalizar(s) {
  return s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}

function plural(n, singular, pluralTxt) {
  const p = pluralTxt || singular + "s";
  return n + " " + (n === 1 ? singular : p);
}

function formatearFecha(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  const hoy = new Date();
  const mismoAnio = d.getFullYear() === hoy.getFullYear();
  const fecha = d.toLocaleDateString("es-ES", mismoAnio
    ? { day: "numeric", month: "short" }
    : { day: "numeric", month: "short", year: "numeric" }).replace(".", "");
  const hora = d.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" });
  return fecha + ", " + hora;
}

function formatearTamano(bytes) {
  if (bytes == null) return "";
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + " KB";
  return (bytes / (1024 * 1024)).toLocaleString("es-ES", { maximumFractionDigits: 1 }) + " MB";
}

function mensajeDe(e) {
  return (e && (e.errorMessage || e.message)) || String(e);
}

let timerAviso = null;
function aviso(texto) {
  const el = $("pie");
  el.textContent = texto;
  clearTimeout(timerAviso);
  timerAviso = setTimeout(() => { if (el.textContent === texto) el.textContent = ""; }, 4000);
}

function mostrarError(texto) { $("estado").textContent = texto; }
function limpiarError() { $("estado").textContent = ""; }

// ---------------------------------------------------------------- vista previa de diseño (sin Outlook)

window.__adjuntosHiloPreview = function (datos) {
  simulado = true;
  modo = datos.compose ? "compose" : "read";
  document.body.classList.toggle("compose", modo === "compose");
  if (!window.__controlesListos) { conectarControles(); window.__controlesListos = true; }
  adjuntosActuales = datos.items;
  nMensajesHilo = datos.nMensajes;
  $("asunto").textContent = datos.asunto || "";
  pintar();
};
