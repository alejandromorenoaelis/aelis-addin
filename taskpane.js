// =====================================================================
//  Aelis · Extraer firma — panel de tareas
//  URL del flujo de Power Automate.
// =====================================================================
const FLOW_URL = "https://default3ec777bd8b8646a8800f6d98eab6bc.39.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7d726e3867224b58a544c874afb6f4be/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=aUXN-WDHSlIjy5RIuXp3tUeXvMvd1fLeluvpG7aWUB4";

// Tamano minimo para considerar que una imagen es una firma y no un icono,
// un separador o un pixel de seguimiento.
const TAMANO_MINIMO = 15360; // 15 KB

// Imagenes que se envian al flujo. Se mandan varias porque el cruce por cid es
// heuristico; el flujo valida cual es la buena comparando dominios y deja de
// hacer OCR en cuanto una valida.
const MAX_IMAGENES = 3;

// Poner en true para ver el detalle de los adjuntos en la consola del panel.
const DEBUG = true;

let datosCorreo = {};

const $ = (id) => document.getElementById(id);

const log = (...args) => { if (DEBUG) console.log("[Aelis firma]", ...args); };

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

  log("Requirement set 1.8 disponible:",
      Office.context.requirements.isSetSupported("Mailbox", "1.8"));
});

// ---------- lectura del correo ----------
function leerCuerpo(formato) {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync(formato, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : "");
    });
  });
}

// Recorta el texto para quedarse solo con el mensaje mas reciente del hilo.
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

// Recorta el HTML por el inicio del mensaje citado.
// OJO: no se corta por el div de firma (id="...Signature..."), porque ese div
// ES la firma del remitente, que es justo lo que queremos conservar.
function cortarHtmlUltimoMensaje(html) {
  const marc = [
    />De:<\/b>/i, />From:<\/b>/i, /border-top:1pt solid/i,
    /id="[^"]*divRplyFwdMsg[^"]*"/i, /class="[^"]*gmail_quote[^"]*"/i
  ];
  let corte = html.length;
  for (const re of marc) { const i = html.search(re); if (i !== -1 && i < corte) corte = i; }
  return html.slice(0, corte);
}

// Devuelve los identificadores cid: referenciados en un fragmento de HTML.
function extraerCids(html) {
  const set = new Set();
  const re = /cid:([^"'\s>&]+)/gi;
  let m;
  while ((m = re.exec(html)) !== null) set.add(m[1].toLowerCase());
  return [...set];
}

function leerContenidoAdjunto(id) {
  return new Promise((resolve) => {
    if (typeof Office.context.mailbox.item.getAttachmentContentAsync !== "function") {
      log("getAttachmentContentAsync no disponible: requiere Mailbox 1.8");
      resolve(null);
      return;
    }
    try {
      Office.context.mailbox.item.getAttachmentContentAsync(id, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded && res.value &&
            res.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          resolve(res.value.content);
        } else {
          log("No se pudo leer el adjunto", id, res.error ? res.error.message : "");
          resolve(null);
        }
      });
    } catch (e) {
      log("Excepcion leyendo el adjunto", id, e.message);
      resolve(null);
    }
  });
}

// Selecciona las imagenes candidatas a ser la firma del remitente.
//
// AttachmentDetails no expone contentId en modo lectura, asi que el cruce se
// hace por la convencion de Outlook: el cid suele empezar por el nombre del
// adjunto (image001.png@01DA...). Se cruza contra el HTML ya recortado al
// ultimo mensaje del hilo, para no coger las firmas de mensajes anteriores.
//
// Si el cruce no encuentra nada (otro cliente de correo, otra convencion de
// nombres) se cae a todas las candidatas ordenadas por peso, y sera el flujo
// quien descarte las que no correspondan al remitente.
function seleccionarImagenesFirma(htmlUltimoMensaje) {
  const adjuntos = Office.context.mailbox.item.attachments || [];
  const cids = extraerCids(htmlUltimoMensaje || "");

  log("Adjuntos recibidos:", adjuntos.length,
      adjuntos.map((a) => ({
        nombre: a.name, tipo: a.contentType, tam: a.size, inline: a.isInline
      })));

  const candidatas = adjuntos.filter((a) =>
    a.isInline === true &&
    (a.contentType || "").indexOf("image/") === 0 &&
    Number(a.size) >= TAMANO_MINIMO
  );

  const delUltimo = candidatas.filter((a) => {
    const base = (a.name || "").toLowerCase();
    return base && cids.some((c) => c.indexOf(base) === 0);
  });

  log("cids del ultimo mensaje:", cids);
  log("Candidatas tras filtrar:", candidatas.length,
      "| del ultimo mensaje:", delUltimo.length,
      delUltimo.length ? "(se usa el cruce por cid)" : "(sin cruce: se usan todas)");

  const seleccion = delUltimo.length ? delUltimo : candidatas;

  return seleccion
    .sort((a, b) => Number(b.size) - Number(a.size))
    .slice(0, MAX_IMAGENES);
}

// ---------- accion principal ----------
async function extraerFirma() {
  const boton = $("run");
  boton.disabled = true;
  boton.innerHTML = '<span class="spin"></span> Analizando\u2026';
  estado("work", "Analizando la firma\u2026", "Leyendo el correo y sus imagenes.");

  try {
    const [texto, html] = await Promise.all([leerCuerpo("text"), leerCuerpo("html")]);
    datosCorreo.cuerpoCompleto = texto;
    datosCorreo.cuerpo = aislarUltimoMensaje(texto);
    datosCorreo.cuerpoHtml = cortarHtmlUltimoMensaje(html);

    const seleccion = seleccionarImagenesFirma(datosCorreo.cuerpoHtml);

    const imagenes = [];
    for (const a of seleccion) {
      const b64 = await leerContenidoAdjunto(a.id);
      if (b64) {
        imagenes.push({ nombre: a.name, tipo: a.contentType, tamano: a.size, base64: b64 });
      }
    }
    datosCorreo.imagenesFirma = imagenes;

    log("Imagenes con contenido leido:", imagenes.length,
        imagenes.map((i) => i.nombre));
    rellenarDetalle(datosCorreo.cuerpo, imagenes);

    if (!FLOW_URL) {
      estado("ok", "Todo listo", "Configura la conexion con Power Automate para enviar.");
      return;
    }

    const res = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify(datosCorreo)
    });

    if (res.ok) {
      exito(imagenes);
    } else {
      const detalle = await res.text().catch(() => "");
      log("El flujo respondio", res.status, detalle);
      estado("err", "No se pudo procesar",
             "El flujo respondio con el codigo " + res.status + ". Revisa el historial de ejecuciones.");
    }
  } catch (e) {
    log("Error en extraerFirma:", e);
    estado("err", "No se pudo enviar",
           (e && e.message ? e.message : "Error de red") + ". Revisa la consola del panel.");
  } finally {
    resetBoton();
  }
}

function exito(imagenes) {
  const sub = imagenes.length
    ? "Firma detectada en imagen \u00b7 " + imagenes.length +
      (imagenes.length === 1 ? " imagen enviada" : " imagenes enviadas")
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
  $("preview").textContent = texto || "(vacio)";
  $("imgs").textContent = imagenes.length
    ? imagenes.map((im) => "\u2022 " + im.nombre + " (" + Math.round(im.tamano / 1024) + " KB)").join("\n")
    : "Sin imagenes (firma en texto).";
}
