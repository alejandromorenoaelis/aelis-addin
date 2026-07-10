// =====================================================================
//  URL del flujo de Power Automate.
// =====================================================================
const FLOW_URL = "https://default3ec777bd8b8646a8800f6d98eab6bc.39.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7d726e3867224b58a544c874afb6f4be/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=aUXN-WDHSlIjy5RIuXp3tUeXvMvd1fLeluvpG7aWUB4";

const TAMANO_MINIMO = 5120; // 5 KB

let datosCorreo = {};

const $ = (id) => document.getElementById(id);

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  datosCorreo = {
    remitente: item.from ? item.from.emailAddress : "",
    nombreRemitente: item.from ? item.from.displayName : "",
    asunto: item.subject || "",
    fecha: item.dateTimeCreated
  };
  $("remitente").textContent = datosCorreo.remitente || "\u2014";
  $("asunto").textContent = datosCorreo.asunto || "\u2014";

  const boton = $("run");
  boton.disabled = false;
  boton.onclick = extraerFirma;
});

// ---------- lectura del correo ----------
function leerCuerpo(formato) {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync(formato, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : "");
    });
  });
}
function aislarUltimoMensaje(texto) {
  const seps = [
    /\r?\nDe:\s/i, /\r?\nFrom:\s/i,
    /-----\s*Mensaje original\s*-----/i, /-----\s*Original Message\s*-----/i,
    /\r?\nEl .*escribio:/i, /\r?\nOn .*wrote:/i, /\r?\n_{5,}/, /\r?\nEnviado desde/i
  ];
  let corte = texto.length;
  for (const re of seps) { const i = texto.search(re); if (i !== -1 && i < corte) corte = i; }
  return texto.slice(0, corte).trim();
}
function cortarHtmlUltimoMensaje(html) {
  const marc = [
    />De:<\/b>/i, />From:<\/b>/i, /border-top:1pt solid/i,
    /id="[^"]*divRplyFwdMsg[^"]*"/i, /class="[^"]*gmail_quote[^"]*"/i, /id="[^"]*x_Signature[^"]*"/i
  ];
  let corte = html.length;
  for (const re of marc) { const i = html.search(re); if (i !== -1 && i < corte) corte = i; }
  return html.slice(0, corte);
}
function extraerCids(html) {
  const set = new Set(); const re = /cid:([^"'\s>&]+)/gi; let m;
  while ((m = re.exec(html)) !== null) set.add(m[1]);
  return [...set];
}
function leerContenidoAdjunto(id) {
  return new Promise((resolve) => {
    try {
      Office.context.mailbox.item.getAttachmentContentAsync(id, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded && res.value &&
            res.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          resolve(res.value.content);
        } else { resolve(null); }
      });
    } catch (e) { resolve(null); }
  });
}

// ---------- acción principal ----------
async function extraerFirma() {
  const boton = $("run");
  boton.disabled = true;
  boton.innerHTML = '<span class="spin"></span> Analizando…';
  estado("work", "Analizando la firma…", "Leyendo el correo y sus imágenes.");

  const [texto, html] = await Promise.all([leerCuerpo("text"), leerCuerpo("html")]);
  datosCorreo.cuerpoCompleto = texto;
  datosCorreo.cuerpo = aislarUltimoMensaje(texto);
  datosCorreo.cuerpoHtml = html;

  // imágenes de la firma del remitente
  const cids = extraerCids(cortarHtmlUltimoMensaje(html));
  const adjuntos = Office.context.mailbox.item.attachments || [];
  const candidatas = adjuntos.filter((a) =>
    a.isInline && (a.contentType || "").indexOf("image/") === 0 &&
    a.contentId && cids.includes(a.contentId) && a.size >= TAMANO_MINIMO
  );
  const imagenes = [];
  for (const a of candidatas) {
    const b64 = await leerContenidoAdjunto(a.id);
    if (b64) imagenes.push({ nombre: a.name, tipo: a.contentType, tamano: a.size, base64: b64 });
  }
  datosCorreo.imagenesFirma = imagenes;

  rellenarDetalle(datosCorreo.cuerpo, imagenes);

  if (!FLOW_URL) {
    estado("ok", "Todo listo", "Configura la conexión con Power Automate para enviar.");
    resetBoton();
    return;
  }

  try {
    const res = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify(datosCorreo)
    });
    if (res.ok) exito(imagenes);
    else estado("err", "No se pudo procesar", "El flujo respondió con el código " + res.status + ". Inténtalo de nuevo.");
  } catch (e) {
    // Puede ser CORS al leer la respuesta; el envío suele haberse completado.
    exito(imagenes);
  }
  resetBoton();
}

function exito(imagenes) {
  const sub = imagenes.length
    ? "Firma detectada en imagen · " + imagenes.length + (imagenes.length === 1 ? " imagen procesada" : " imágenes procesadas")
    : "Firma detectada en el texto del correo.";
  estado("ok", "Firma enviada", sub);
}

function resetBoton() {
  const boton = $("run");
  boton.disabled = false;
  boton.textContent = "Extraer firma";
}

// ---------- UI ----------
function estado(tipo, msg, sub) {
  const box = $("status");
  box.hidden = false;
  box.className = "status " + tipo;
  const ico = tipo === "ok"
    ? '<svg class="check" viewBox="0 0 24 24"><circle cx="12" cy="12" r="12"/><path d="M6.5 12.5l3.5 3.5 7.5-8"/></svg>'
    : tipo === "err"
    ? '<span class="ico" style="color:var(--danger);font-weight:800;">!</span>'
    : '<span class="spin"></span>';
  box.innerHTML = '<span class="ico">' + ico + '</span><div><div class="msg">' +
    msg + '</div><div class="sub">' + (sub || "") + '</div></div>';
}

function rellenarDetalle(texto, imagenes) {
  $("detalle").hidden = false;
  $("preview").textContent = texto || "(vacío)";
  $("imgs").textContent = imagenes.length
    ? imagenes.map((im) => "• " + im.nombre + " (" + Math.round(im.tamano / 1024) + " KB)").join("\n")
    : "Sin imágenes (firma en texto).";
}
