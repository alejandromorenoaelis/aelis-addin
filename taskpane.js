// =====================================================================
//  URL del flujo de Power Automate (disparador "Cuando se recibe una
//  solicitud HTTP"). Pegala entre las comillas tras crear el flujo.
// =====================================================================
const FLOW_URL = "";

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

// Lee el cuerpo del correo en texto plano.
function leerCuerpo() {
  return new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync("text", (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : "");
    });
  });
}

// Corta las cadenas de respuesta y se queda con el ultimo mensaje,
// que es donde normalmente esta la firma del remitente.
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

async function extraerFirma() {
  const status = document.getElementById("status");

  status.textContent = "Leyendo el correo...";
  const completo = await leerCuerpo();
  datosCorreo.cuerpoCompleto = completo;
  datosCorreo.cuerpo = aislarUltimoMensaje(completo);
  mostrarPreview(datosCorreo.cuerpo);

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
    pre.style.cssText = "background:#f3f2f1;border-radius:4px;padding:8px;font-size:12px;white-space:pre-wrap;max-height:160px;overflow:auto;margin:2px 0 12px;";
    const boton = document.getElementById("run");
    boton.parentNode.insertBefore(label, boton);
    boton.parentNode.insertBefore(pre, boton);
  }
  pre.textContent = texto || "(vacio)";
}
