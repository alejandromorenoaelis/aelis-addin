// =====================================================================
//  FASE 2 — Power Automate
//  Pega aquí la URL del disparador "Cuando se recibe una solicitud HTTP"
//  de tu flujo. Mientras esté vacía, el botón solo confirma que funciona.
// =====================================================================
const FLOW_URL = "";

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  document.getElementById("subject").textContent = item.subject || "(sin asunto)";
  document.getElementById("run").onclick = ejecutarProceso;
});

async function ejecutarProceso() {
  const status = document.getElementById("status");
  const item = Office.context.mailbox.item;

  // Fase 1: sin flujo configurado todavía -> solo confirmamos que el add-in va bien.
  if (!FLOW_URL) {
    status.textContent = "✅ Add-in funcionando. Configura FLOW_URL en taskpane.js para lanzar el flujo.";
    return;
  }

  // Fase 2: enviamos datos del correo a Power Automate.
  status.textContent = "Enviando a Power Automate…";
  try {
    const res = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        asunto: item.subject,
        remitente: item.from ? item.from.emailAddress : null,
        fechaCreacion: item.dateTimeCreated,
        idConversacion: item.conversationId
      })
    });
    status.textContent = res.ok
      ? "✅ Proceso lanzado correctamente."
      : "⚠️ El flujo respondió con error " + res.status + ".";
  } catch (e) {
    // Si ves un error de CORS aquí, añade una acción "Respuesta" en el flujo
    // devolviendo la cabecera  Access-Control-Allow-Origin: *
    status.textContent = "❌ Error de red o CORS: " + e.message;
  }
}
