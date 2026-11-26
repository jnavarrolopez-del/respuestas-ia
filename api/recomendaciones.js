// Función para Vercel: POST /api/recomendaciones (misma lógica que el servidor local).
// Configura ANTHROPIC_API_KEY en Vercel → Project → Settings → Environment Variables.
const { validarAgregados, pedirRecomendaciones, respuestaDeError, crearLimitador } = require("../server/recomendaciones");

const permitir = crearLimitador(10);

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  if (req.method !== "POST") return res.status(405).json({ error: "Usa POST." });
  const ip = String(req.headers["x-forwarded-for"] || "").split(",")[0].trim() || "desconocida";
  if (!permitir(ip)) return res.status(429).json({ error: "Demasiadas peticiones. Prueba en un minuto." });
  try {
    const cuerpo = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    if (JSON.stringify(cuerpo || {}).length > 20 * 1024) return res.status(413).json({ error: "Datos demasiado grandes." });
    const recomendaciones = await pedirRecomendaciones(validarAgregados(cuerpo));
    return res.status(200).json({ recomendaciones });
  } catch (err) {
    if (err instanceof SyntaxError) return res.status(400).json({ error: "El cuerpo no es JSON válido." });
    const { status, mensaje } = respuestaDeError(err);
    console.warn("[recomendaciones]", status, err && err.message);
    return res.status(status).json({ error: mensaje });
  }
};
