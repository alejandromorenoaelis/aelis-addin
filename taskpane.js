// =====================================================================
//  URL del flujo de Power Automate.
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

function leerCuerpo(formato) {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync(formato, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : "");
    });
  });
}

function aislarUltimoMensaje(texto) {
  const separadores = [
    /\r?\nDe:\s/i, /\r?\nFrom:\s/i,
    /-----\s*Mensaje original\s*-----/i, /-----\s*Original Message\s*-----/i,
    /\r?\nEl .*escribio:/i, /\r?\nOn .*wrote:/i,
    /\r?\n_{5,}/, /\r?\nEnviado desde/i
  ];
  let corte = texto.length;
  for (const re of separadores) {
    const i = texto.search(re);
    if (i !== -1 && i < corte) corte = i;
  }
  return texto.slice(0, corte).trim();
}

// Corta el HTML en el primer separador de respuesta -> queda el ultimo mensaje.
function cortarHtmlUltimoMensaje(html) {
  const marcadores = [
    />De:<\/b>/i, />From:<\/b>/i,
    /border-top:1pt solid/i,
    /id="[^"]*divRplyFwdMsg[^"]*"/i,
    /class="[^"]*gmail_quote[^"]*"/i,
    /id="[^"]*x_Signature[^"]*"/i
  ];
  let corte = html.length;
  for (const re of marcadores) {
    const i = html.search(re);
    if (i !== -1 && i < corte) corte = i;
  }
  return html.slice(0, corte);
}

// Extrae los valores cid: que aparecen en un trozo de HTML.
function extraerCids(html) {
  const set = new Set();
  const re = /cid:([^"'\s>&]+)/gi;
  let m;
  while ((m = re.exec(html)) !== null) set.add(m[1]);
  return [...set];
}

function listarAdjuntos() {
  const lista = Office.context.mailbox.item.attachments || [];
  return lista.map((a) => ({
    nombre: a.name, tipo: a.contentType, tamano: a.size,
    inline: a.isInline, contentId: a.contentId, id: a.id
  }));
}

async function extraerFirma() {
  const status = document.getElementById("status");
  status.textContent = "Leyendo el correo...";

  const [textoCompleto, htmlCompleto] = await Promise.all([
    leerCuerpo("text"), leerCuerpo("html")
  ]);

  datosCorreo.cuerpoCompleto = textoCompleto;
  datosCorreo.cuerpo = aislarUltimoMensaje(textoCompleto);
  datosCorreo.cuerpoHtml = htmlCompleto;

  const adjuntos = listarAdjuntos();
  const htmlTop = cortarHtmlUltimoMensaje(htmlCompleto);
  const cidsUltimo = extraerCids(htmlTop);

  mostrarPreview(datosCorreo.cuerpo);
  mostrarDiagnostico(adjuntos, cidsUltimo);

  if (!FLOW_URL) {
    status.textContent = "\u2705 Add-in funcionando (modo diagnostico).";
    return;
  }
  status.textContent = "Enviando a Power Automate...";
  try {
    const res = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify(datosCorreo)
    });
    status.textContent = res.ok ? "\u2705 Enviado a Power Automate."
      : "\u26A0\uFE0F El flujo respondio con codigo " + res.status + ".";
  } catch (e) {
    status.textContent = "\u2705 Enviado (revisa el flujo).";
  }
}

function mostrarPreview(texto) {
  let pre = document.getElementById("preview");
  if (!pre) {
    const label = document.createElement("div");
    label.className = "campo";
    label.textContent = "Texto detectado (ultimo mensaje):";
    pre = document.createElement("pre");
    pre.id = "preview";
    pre.style.cssText = "background:#f3f2f1;border-radius:4px;padding:8px;font-size:12px;white-space:pre-wrap;max-height:120px;overflow:auto;margin:2px 0 12px;";
    const boton = document.getElementById("run");
    boton.parentNode.insertBefore(label, boton);
    boton.parentNode.insertBefore(pre, boton);
  }
  pre.textContent = texto || "(vacio)";
}

function mostrarDiagnostico(adjuntos, cidsUltimo) {
  let box = document.getElementById("diag");
  if (!box) {
    const label = document.createElement("div");
    label.className = "campo";
    label.textContent = "DIAGNOSTICO firma-imagen:";
    box = document.createElement("pre");
    box.id = "diag";
    box.style.cssText = "background:#fff4ce;border:1px solid #ffd666;border-radius:4px;padding:8px;font-size:11px;white-space:pre-wrap;max-height:260px;overflow:auto;margin:2px 0 12px;";
    const boton = document.getElementById("run");
    boton.parentNode.insertBefore(label, boton);
    boton.parentNode.insertBefore(box, boton);
  }

  let txt = "== cids en el ULTIMO mensaje (HTML) ==\n";
  txt += cidsUltimo.length ? cidsUltimo.join("\n") : "(ninguno)";
  txt += "\n\n== Adjuntos / imagenes ==\n";
  if (!adjuntos.length) {
    txt += "(no se detectaron)";
  } else {
    adjuntos.forEach((a, i) => {
      const casa = a.contentId && cidsUltimo.includes(a.contentId);
      txt += (i + 1) + ") " + (a.nombre || "(sin nombre)") +
        "\n   tipo: " + a.tipo +
        "\n   tamano: " + a.tamano + " bytes" +
        "\n   inline: " + a.inline +
        "\n   contentId: " + (a.contentId || "(NO disponible)") +
        "\n   >>> " + (casa ? "COINCIDE con el ultimo mensaje" : "no coincide / sin dato") +
        "\n\n";
    });
  }
  box.textContent = txt;
}
