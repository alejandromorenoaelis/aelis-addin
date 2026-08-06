// =====================================================================
//  Aelis · Extraer firma — panel de tareas
//  URL del flujo de Power Automate.
// =====================================================================
const FLOW_URL = "https://default3ec777bd8b8646a8800f6d98eab6bc.39.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7d726e3867224b58a544c874afb6f4be/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=aUXN-WDHSlIjy5RIuXp3tUeXvMvd1fLeluvpG7aWUB4";

// Tamano minimo para considerar que una imagen es una firma y no un icono,
// un separador o un pixel de seguimiento.
const TAMANO_MINIMO = 15360; // 15 KB

// Imagenes que se envian al flujo. Se mandan varias porque el cruce por cid es
// heuristico; el flujo valida cual es la buena comparando dominios.
const MAX_IMAGENES = 3;

// Poner en true para ver el detalle en la consola del panel.
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
    fecha: item.dateTimeCreated,
    // Email del usuario del buzon. Permite al flujo decidir, segun despliegue,
    // si se descartan las firmas del propio dominio del usuario.
    usuarioActual: (Office.context.mailbox.userProfile &&
                    Office.context.mailbox.userProfile.emailAddress) || ""
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

// Aplica una lista de patrones y devuelve el texto cortado por el primero
// que aparezca. Si ninguno casa, devuelve el texto intacto.
function cortarPorPrimerPatron(texto, patrones) {
  let corte = texto.length;
  for (const re of patrones) {
    const i = texto.search(re);
    if (i !== -1 && i < corte) corte = i;
  }
  return texto.slice(0, corte);
}

// Recorta el texto para quedarse solo con el mensaje mas reciente del hilo.
//
// IMPORTANTE: los patrones NO exigen salto de linea antes del encabezado.
// Al convertir el HTML a texto plano, Outlook web pega el "De:" al final del
// parrafo anterior ("...puedan derivarse De: Alejandro Moreno <...>"), asi que
// anclar en \n hacia que el corte no se aplicase nunca.
function aislarUltimoMensaje(texto) {
  const seps = [
    // Encabezado citado con email entre angulos: "De: Nombre <a@b.com>"
    /\s(?:De|From|Von|Da|De la part de):\s*[^<\n]{0,80}<[^>@\s]+@[^>\s]+>/i,
    // Encabezado citado sin angulos, seguido de "Enviado el" / "Sent"
    /\s(?:De|From):\s.{0,80}?\s(?:Enviado el|Enviado|Sent|Gesendet):/i,
    // Separadores clasicos
    /-----\s*Mensaje original\s*-----/i,
    /-----\s*Original Message\s*-----/i,
    // Gmail y Apple Mail
    /\s?El\s.{0,140}?escribi[oó]:/i,
    /\s?On\s.{0,140}?wrote:/i,
    // Otros
    /\r?\n_{5,}/,
    /\r?\nEnviado desde/i,
    /\r?\nObtener Outlook para/i
  ];
  return cortarPorPrimerPatron(texto, seps).trim();
}

// Quita los avisos legales y coletillas del pie. Aportan ruido (razones
// sociales distintas, buzones de proteccion de datos) que despista al modelo.
function quitarAvisoLegal(texto) {
  const marc = [
    /De conformidad con el Reglamento/i,
    /En cumplimiento de lo establecido/i,
    /Colaboremos con el medio ambiente/i,
    /Este mensaje y cualquier documento/i,
    /Este correo electr[oó]nico y sus/i,
    /La informaci[oó]n contenida en este/i,
    /Antes de imprimir este/i,
    /This e?-?mail (?:and any|message) /i,
    /The information contained in this/i
  ];
  return cortarPorPrimerPatron(texto, marc).trim();
}

// Recorta el HTML por el inicio del mensaje citado.
// OJO: no se corta por el div de firma (id="...Signature..."), porque ese div
// ES la firma del remitente, que es justo lo que queremos conservar.
function cortarHtmlUltimoMensaje(html) {
  const marc = [
    /id="[^"]*divRplyFwdMsg[^"]*"/i,
    /class="[^"]*gmail_quote[^"]*"/i,
    /<blockquote[^>]*type="cite"/i,
    />\s*(?:De|From):\s*<\/b>/i,
    /border-top:\s?1pt solid/i
  ];
  return cortarPorPrimerPatron(html, marc);
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
// ultimo mensaje. Si el cruce no da nada, se usan todas las candidatas y sera
// el flujo quien descarte las que no correspondan al remitente.
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
  log("Candidatas:", candidatas.length, "| del ultimo mensaje:", delUltimo.length,
      delUltimo.length ? "(cruce por cid aplicado)" : "(sin cruce: se usan todas)");

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
    datosCorreo.cuerpo = quitarAvisoLegal(aislarUltimoMensaje(texto));
    datosCorreo.cuerpoHtml = cortarHtmlUltimoMensaje(html);

    log("Longitud cuerpo original:", texto.length,
        "| tras recortar:", datosCorreo.cuerpo.length);
    if (datosCorreo.cuerpo.length === texto.length) {
      log("AVISO: el recorte no encontro ningun separador. El cuerpo puede " +
          "contener mensajes citados. La validacion por dominio del flujo es " +
          "la que debe filtrar en este caso.");
    }

    const seleccion = seleccionarImagenesFirma(datosCorreo.cuerpoHtml);

    const imagenes = [];
    for (const a of seleccion) {
      const b64 = await leerContenidoAdjunto(a.id);
      if (b64) {
        imagenes.push({ nombre: a.name, tipo: a.contentType, tamano: a.size, base64: b64 });
      }
    }
    datosCorreo.imagenesFirma = imagenes;

    log("Imagenes con contenido leido:", imagenes.length, imagenes.map((i) => i.nombre));
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
