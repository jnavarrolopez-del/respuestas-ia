/*
 * Punto de entrada del runtime compartido.
 * Al abrir el libro se carga en segundo plano (sin panel) y registra los eventos.
 * Si el usuario abre el panel desde "Configurar clave", muestra el formulario.
 */
import { registerHandlers, readKey, writeKey, makeClient, MODEL } from "./addin.js";

const form = document.getElementById("form");
const input = document.getElementById("key");
const save = document.getElementById("save");
const statusEl = document.getElementById("status");

function setStatus(text, kind) {
  statusEl.textContent = text;
  statusEl.className = kind || "";
}

if (readKey()) {
  input.placeholder = "Clave guardada (escribe otra para cambiarla)";
}

input.addEventListener("input", () => {
  save.disabled = input.value.trim().length < 10;
  setStatus("");
});

form.addEventListener("submit", async (ev) => {
  ev.preventDefault();
  const key = input.value.trim();
  if (!key) return;
  save.disabled = true;
  setStatus("Comprobando la clave…");
  try {
    // Comprobación gratuita: consultar el modelo no consume tokens.
    await makeClient(key).models.retrieve(MODEL);
  } catch (err) {
    const status = err && err.status;
    setStatus(
      status === 401 || status === 403
        ? "La clave no es válida. Revísala y vuelve a intentarlo."
        : "No se pudo conectar para comprobar la clave. Revisa la conexión y vuelve a intentarlo.",
      "error"
    );
    save.disabled = false;
    return;
  }

  try {
    writeKey(key);
  } catch {
    setStatus("No se pudo guardar la clave en este equipo.", "error");
    save.disabled = false;
    return;
  }
  input.value = "";
  input.placeholder = "Clave guardada (escribe otra para cambiarla)";

  try {
    // A partir de ahora el complemento se carga solo, sin panel, al abrir un libro.
    await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  } catch (err) {
    console.warn("[Respuestas IA] setStartupBehavior:", err);
  }
  setStatus("Listo. Ya puedes cerrar este panel.", "ok");
  setTimeout(() => {
    Office.addin.hide().catch(() => {});
  }, 2500);
});

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;
  try {
    await registerHandlers();
  } catch (err) {
    console.warn("[Respuestas IA] No se pudieron registrar los eventos:", err);
  }
});
