# Aelis Procesos — add-in de Outlook

Botón en el ribbon de Outlook que abre un panel y (fase 2) llama a un flujo de Power Automate.

## Paso 1 — Editar la URL base
En `manifest.xml`, busca y reemplaza **todas** las apariciones de:

    https://tu-usuario.github.io/aelis-addin

por la URL real donde vayas a alojar estos archivos.

## Paso 2 — Publicar los archivos (HTTPS obligatorio)
La forma gratis más rápida es **GitHub Pages**:

1. Crea un repo (p. ej. `aelis-addin`) y sube esta carpeta (sin el `manifest.xml` si quieres, aunque no molesta).
2. Settings → Pages → Deploy from branch → `main` / raíz.
3. En 1-2 min tendrás: `https://TU-USUARIO.github.io/aelis-addin/taskpane.html`
4. Esa base (`https://TU-USUARIO.github.io/aelis-addin`) es la que va en el manifiesto (Paso 1).

Alternativas válidas: Azure Static Web Apps, cualquier hosting con HTTPS y certificado válido.

## Paso 3 — Cargar el add-in en Outlook (sideload, solo para ti)
En **Outlook en la web / nuevo Outlook**:

1. Abre un correo → menú **… (Más acciones)** → **Complementos** / **Obtener complementos**.
2. **Mis complementos** → **Complementos personalizados** → **Agregar un complemento personalizado** → **Agregar desde archivo**.
3. Selecciona tu `manifest.xml`. Acepta el aviso de que no viene de la tienda.
4. Abre cualquier correo: verás el botón **Procesos** (grupo *Aelis*) en el ribbon o en el menú **…**.

> Para desplegarlo a toda la empresa hace falta el Centro de administración de M365.
> (Apps integradas / implementación centralizada) → requiere un administrador global de Aelis.

## Fase 2 — Conectar Power Automate
1. Crea un flujo con el disparador **"Cuando se recibe una solicitud HTTP"** (conector premium).
2. Copia la URL POST que genera.
3. Pégala en `FLOW_URL` dentro de `taskpane.js` y vuelve a publicar.
4. El panel enviará asunto, remitente y fecha del correo al flujo.

Si el `fetch` falla por **CORS**, añade una acción **"Respuesta"** al flujo que devuelva
la cabecera `Access-Control-Allow-Origin: *`.
