// =====================================================================
//  Pega aquí la URL del disparador HTTP de tu flujo de Power Automate.
//  (La consigues al crear el flujo, en el paso "Cuando se recibe una
//   solicitud HTTP", tras guardar.)
// =====================================================================
const FLOW_URL = "";

// Configuración de las 4 opciones del menú.
const ACCIONES = {
  registrar: {
    titulo: "Registrar correo en seguimiento",
    descripcion: "Guarda el remitente, asunto y fecha en la tabla de seguimiento.",
    activa: true
  },
  aprobacion: {
    titulo: "Enviar a aprobación",
    descripcion: "Esta acción se activará en la siguiente fase del proyecto.",
    activa: false
  },
  adjuntos: {
    titulo: "Guardar adjuntos en SharePoint",
    descripcion: "Esta acción se activará en la siguiente fase del proyecto.",
    activa: false
  },
  documento: {
    titulo: "Generar documento desde plantilla",
    descripcion: "Esta acción se activará en la siguiente fase del proyecto.",
    activa: false
  }
};

let datosCorreo = {};

Office.onReady(() => {
  const accion = new URLSearchParams(location.search).get("action") || "registrar";
  const cfg = ACCIONES[accion] || ACCIONES.registrar;

  const item = Office.context.mailbox.item;
  datosCorreo = {
    accion: accion,
    remitente: item.from ? item.from.emailAddress : "",
    asunto: item.subject || "",
    fecha: item.dateTimeCreated
  };

  document.getElementById("titulo").textContent = cfg.titulo;
  document.getElementById("descripcion").textContent = cfg.descripcion;
  document.getElementById("remitente").textContent = datosCorreo.remitente || "—";
  document.getElementById("asunto").textContent = datosCorreo.asunto || "—";

  const boton = document.getElementById("run");
  if (cfg.activa) {
    boton.textContent = "Ejecutar";
    boton.disabled = false;
    boton.onclick = ejecutar;
  } else {
    boton.textContent = "Próximamente (fase 2)";
    boton.disabled = true;
  }
});

async function ejecutar() {
  const status = document.getElementById("status");

  if (!FLOW_URL) {
    status.textContent = "✅ Add-in funcionando. Falta configurar FLOW_URL en taskpane.js.";
    return;
  }

  status.textContent = "Enviando a Power Automate…";
  try {
    const res = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify(datosCorreo)
    });
    status.textContent = res.ok
      ? "✅ Registrado correctamente en la tabla de seguimiento."
      : "⚠️ El flujo respondió con código " + res.status + ".";
  } catch (e) {
    // Un error aquí suele ser de CORS al leer la respuesta; el registro
    // normalmente se ha creado igualmente. Comprueba el Excel.
    status.textContent = "✅ Enviado. Comprueba la fila nueva en el Excel de seguimiento.";
  }
}
