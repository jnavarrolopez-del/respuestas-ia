/*
 * Recomendaciones para el dueño del taller, generadas por Claude EN EL SERVIDOR.
 * La clave de la API (ANTHROPIC_API_KEY) solo existe aquí, nunca en el navegador.
 * Solo se aceptan cifras agregadas con una forma fija: nada de reservas sueltas ni matrículas.
 */
const sdk = require("@anthropic-ai/sdk");

const Anthropic = sdk.default || sdk;
const MODELO = "claude-opus-5-5";
const DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
const PORCENTAJES = ["aceptadas", "reprogramadas", "canceladas", "rechazadas", "no_presentadas", "pendientes"];

class ErrorDatos extends Error {}

function entero(v, max = 100000) {
  if (!Number.isInteger(v) || v < 0 || v > max) throw new ErrorDatos("número no válido");
  return v;
}
function numero(v, max = 100000) {
  if (typeof v !== "number" || !Number.isFinite(v) || v < 0 || v > max) throw new ErrorDatos("número no válido");
  return Math.round(v * 10) / 10;
}
function fecha(v) {
  if (v === null) return null;
  if (typeof v !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new ErrorDatos("fecha no válida");
  return v;
}
function lista(v, max, fn) {
  if (!Array.isArray(v) || v.length > max) throw new ErrorDatos("lista no válida");
  return v.map(fn);
}
/** Nombre de servicio: texto corto, sin caracteres de control. */
function nombreServicio(v) {
  if (typeof v !== "string") throw new ErrorDatos("servicio no válido");
  const t = v.replace(/[\u0000-\u001f\u007f]/g, " ").replace(/\s+/g, " ").trim().slice(0, 40);
  if (!t) throw new ErrorDatos("servicio no válido");
  return t;
}

/** Comprueba y reconstruye el cuerpo de la petición. Lanza ErrorDatos si no tiene la forma esperada. */
function validarAgregados(body) {
  if (!body || typeof body !== "object" || Array.isArray(body)) throw new ErrorDatos("cuerpo no válido");
  const p = body.porcentajes;
  if (!p || typeof p !== "object") throw new ErrorDatos("faltan los porcentajes");
  return {
    total: entero(body.total),
    desde: fecha(body.desde ?? null),
    hasta: fecha(body.hasta ?? null),
    mediaPorDia: numero(body.mediaPorDia),
    diasSemana: lista(body.diasSemana, 7, (d) => {
      if (!d || !DIAS.includes(d.dia)) throw new ErrorDatos("día no válido");
      return { dia: d.dia, reservas: entero(d.reservas) };
    }),
    franjasMasLlenas: lista(body.franjasMasLlenas, 5, franja),
    franjasMasVacias: lista(body.franjasMasVacias, 5, franja),
    servicios: lista(body.servicios, 12, (s) => ({
      servicio: nombreServicio(s && s.servicio),
      reservas: entero(s && s.reservas),
      porcentaje: numero(s && s.porcentaje, 100),
    })),
    porcentajes: Object.fromEntries(PORCENTAJES.map((k) => [k, numero(p[k] ?? 0, 100)])),
  };
}
function franja(f) {
  if (!f || typeof f.franja !== "string" || !/^\d{2}:\d{2}$/.test(f.franja)) throw new ErrorDatos("franja no válida");
  return { franja: f.franja, reservas: entero(f.reservas) };
}

const SISTEMA = [
  "Eres un asesor de talleres mecánicos en España. Recibes las cifras agregadas de las reservas de un taller (no ves reservas sueltas ni datos de clientes).",
  "Da entre 2 y 3 recomendaciones para el dueño del taller:",
  "- En español de España, tuteando.",
  "- Cortas: una o dos frases, como mucho 35 palabras cada una.",
  "- Concretas y accionables, apoyadas en las cifras (cita la cifra que la justifica).",
  "- Sin saludos, sin introducción y sin repetir los datos.",
].join("\n");

const ESQUEMA = {
  type: "object",
  properties: { recomendaciones: { type: "array", items: { type: "string" } } },
  required: ["recomendaciones"],
  additionalProperties: false,
};

/** Pide a Claude las recomendaciones. Devuelve un array de 2 o 3 textos. */
async function pedirRecomendaciones(agregados, opciones = {}) {
  const apiKey = opciones.apiKey || process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("Falta ANTHROPIC_API_KEY en el servidor");
  const client = opciones.client || new Anthropic({ apiKey, timeout: 60_000, maxRetries: 2 });
  const respuesta = await client.beta.messages.create({
    model: MODELO,
    max_tokens: 16000,
    betas: ["server-side-fallback-2026-07-01"],
    fallbacks: "default", // si el modelo rechaza la petición, el servidor de Anthropic prueba con otro
    output_config: { effort: "low", format: { type: "json_schema", schema: ESQUEMA } },
    system: SISTEMA,
    messages: [{ role: "user", content: "Cifras de reservas del taller:\n" + JSON.stringify(agregados) }],
  });
  if (respuesta.stop_reason === "refusal" || respuesta.stop_reason === "max_tokens") {
    throw new Error(`Respuesta incompleta de la IA (${respuesta.stop_reason})`);
  }
  const texto = respuesta.content.filter((b) => b.type === "text").map((b) => b.text).join("");
  const lista = JSON.parse(texto).recomendaciones;
  const limpias = (Array.isArray(lista) ? lista : [])
    .filter((r) => typeof r === "string" && r.trim())
    .map((r) => r.trim())
    .slice(0, 3);
  if (limpias.length < 2) throw new Error("La IA no devolvió recomendaciones suficientes");
  return limpias;
}

/** Traduce un error a una respuesta HTTP con un mensaje para el panel (sin detalles internos). */
function respuestaDeError(err) {
  if (err instanceof ErrorDatos) return { status: 400, mensaje: "Los datos enviados no tienen el formato esperado." };
  if (err instanceof Anthropic.RateLimitError) return { status: 429, mensaje: "Hay demasiadas peticiones a la IA. Prueba en un minuto." };
  if (err instanceof Anthropic.AuthenticationError) return { status: 500, mensaje: "La clave de la IA del servidor no es válida." };
  if (err instanceof Anthropic.APIConnectionError) return { status: 502, mensaje: "El servidor no pudo conectar con la IA." };
  if (err instanceof Anthropic.APIError) return { status: 502, mensaje: "La IA no pudo responder ahora." };
  if (/ANTHROPIC_API_KEY/.test(String(err && err.message))) return { status: 500, mensaje: "El servidor no tiene configurada la clave de la IA." };
  return { status: 502, mensaje: "No se pudieron generar las recomendaciones." };
}

/** Límite sencillo por IP: como mucho `max` peticiones por minuto. */
function crearLimitador(max = 10) {
  const visitas = new Map();
  return (ip) => {
    const ahora = Date.now();
    const recientes = (visitas.get(ip) || []).filter((t) => ahora - t < 60_000);
    if (recientes.length >= max) return false;
    recientes.push(ahora);
    visitas.set(ip, recientes);
    return true;
  };
}

module.exports = { validarAgregados, pedirRecomendaciones, respuestaDeError, crearLimitador, ErrorDatos, SISTEMA, MODELO };
