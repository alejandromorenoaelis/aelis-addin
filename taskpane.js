// =====================================================================
//  URL del flujo de Power Automate (disparador "Cuando se recibe una
//  solicitud HTTP").
// =====================================================================
const FLOW_URL = "https://default3ec777bd8b8646a8800f6d98eab6bc.39.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7d726e3867224b58a544c874afb6f4be/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=aUXN-WDHSlIjy5RIuXp3tUeXvMvd1fLeluvpG7aWUB4";

let datosCorreo = {};

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  datosCorreo = {
    remitente: item.from ? item.from.emailAddress : "",
    nombreRemitente: item.from ? item.from.displayName : "",
    asunto: item.subject || "",
    fecha: item.dateTimeCreated
  };

  document.getElementById("titulo").textContent = "Extraer firma";
  document.getElementById("descripcion").textContent =
    "Lee el cuerpo del correo y lo envia a Power Automate para extraer la firma.";
  document.getElementById("remitente").textContent = datosCorreo.remitente || "\u2014";
  document.getElementById("asunto").textContent = datosCorreo.asunto || "\u2014";

  const boton = document.getElementById("run");
  boton.textContent = "Extraer firma";
  boton.disabled = false;
  boton.onclick = extraerFirma;
});

// Lee el cuerpo del correo en el formato indicado ("text" o "html").
function leerCuerpo(formato) {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync(formato, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : "");
    });
  });
}

// Corta las cadenas de respuesta y se queda con el ultimo mensaje.
function aislarUltimoMensaje(texto) {
  const separadores = [
    /\r?\nDe:\s/i,
    /\r?\nFrom:\s/i,
    /-----\s*Mensaje original\s*-----/i,
    /-----\s*Original Message\s*-----/i,
    /\r?\nEl .*escribio:/i,
    /\r?\nOn .*wrote:/i,
    /\r?\n_{5,}/,
    /\r?\nEnviado desde/i
  ];
  let corte = texto.length;
  for (const re of separadores) {
    const i = texto.search(re);
    if (i !== -1 && i < corte) corte = i;
  }
  return texto.slice(0, corte).trim();
}

// --- NUEVO (diagnostico): lista los adjuntos/imagenes que ve en el correo.
function listarAdjuntos() {
  const item = Office.context.mailbox.item;
  const lista = item.attachments || [];
  return lista.map((a) => ({
    nombre: a.name,
    tipo: a.contentType,
    tamano: a.size,
    inline: a.isInline,
    clase: a.attachmentType,
    id: a.id
  }));
}

async function extraerFirma() {
  const status = document.getElementById("status");

  status.textContent = "Leyendo el correo...";

  const [textoCompleto, htmlCompleto] = await Promise.all([
    leerCuerpo("text"),
    leerCuerpo("html")
  ]);

  datosCorreo.cuerpoCompleto = textoCompleto;
  datosCorreo.cuerpo = aislarUltimoMensaje(textoCompleto);
  datosCorreo.cuerpoHtml = htmlCompleto;

  // Diagnostico de adjuntos/imagenes
  const adjuntos = listarAdjuntos();
  datosCorreo.adjuntos = adjuntos;

  mostrarPreview(datosCorreo.cuerpo);
  mostrarAdjuntos(adjuntos);

  if (!FLOW_URL) {
    status.textContent = "\u2705 Add-in funcionando. Falta configurar FLOW_URL en taskpane.js.";
    return;
  }

  status.textContent = "Enviando a Power Automate...";
  try {
    const res = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify(datosCorreo)
    });
    status.textContent = res.ok
      ? "\u2705 Enviado correctamente a Power Automate."
      : "\u26A0\uFE0F El flujo respondio con codigo " + res.status + ".";
  } catch (e) {
    status.textContent = "\u2705 Enviado. Si no ves confirmacion, revisa el flujo en Power Automate.";
  }
}

// Muestra en el panel el texto detectado.
function mostrarPreview(texto) {
  let pre = document.getElementById("preview");
  if (!pre) {
    const label = document.createElement("div");
    label.className = "campo";
    label.textContent = "Texto detectado (ultimo mensaje):";
    pre = document.createElement("pre");
    pre.id = "preview";
    pre.style.cssText = "background:#f3f2f1;border-radius:4px;padding:8px;font-size:12px;white-space:pre-wrap;max-height:140px;overflow:auto;margin:2px 0 12px;";
    const boton = document.getElementById("run");
    boton.parentNode.insertBefore(label, boton);
    boton.parentNode.insertBefore(pre, boton);
  }
  pre.textContent = texto || "(vacio)";
}

// --- NUEVO (diagnostico): muestra en el panel la lista de adjuntos detectados.
function mostrarAdjuntos(adjuntos) {
  let box = document.getElementById("adjuntos");
  if (!box) {
    const label = document.createElement("div");
    label.className = "campo";
    label.textContent = "Adjuntos / imagenes detectados:";
    box = document.createElement("pre");
    box.id = "adjuntos";
    box.style.cssText = "background:#fff4ce;border:1px solid #ffd666;border-radius:4px;padding:8px;font-size:11px;white-space:pre-wrap;max-height:200px;overflow:auto;margin:2px 0 12px;";
    const boton = document.getElementById("run");
    boton.parentNode.insertBefore(label, boton);
    boton.parentNode.insertBefore(box, boton);
  }
  if (!adjuntos.length) {
    box.textContent = "(no se detectaron adjuntos ni imagenes)";
    return;
  }
  box.textContent = adjuntos.map((a, i) =>
    (i + 1) + ") " + (a.nombre || "(sin nombre)") +
    "\n   tipo: " + a.tipo +
    "\n   inline: " + a.inline +
    "\n   clase: " + a.clase +
    "\n   tamano: " + a.tamano + " bytes"
  ).join("\n\n");
}