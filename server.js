// Servidor HTTPS local para el complemento en https://localhost:3000
// - Sirve los archivos de src/ (el panel).
// - POST /api/recomendaciones: pide a Claude las recomendaciones con la clave del archivo .env.
//   La clave nunca llega al navegador.
// Certificados de desarrollo de Office (una vez):  npx office-addin-dev-certs install
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

cargarEnv(path.join(__dirname, ".env"));
const { validarAgregados, pedirRecomendaciones, respuestaDeError, crearLimitador } = require("./server/recomendaciones");

const PORT = Number(process.env.PORT) || 3000;
const ROOT = path.join(__dirname, "src");
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const MAX_CUERPO = 20 * 1024;
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".png": "image/png",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
};
const permitir = crearLimitador(10);

/** Lee un .env sencillo (CLAVE=valor) sin dependencias. No pisa variables ya definidas. */
function cargarEnv(archivo) {
  let texto;
  try {
    texto = fs.readFileSync(archivo, "utf8");
  } catch {
    return;
  }
  for (const linea of texto.split(/\r?\n/)) {
    const m = linea.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/);
    if (m && !(m[1] in process.env)) process.env[m[1]] = m[2].replace(/^["']|["']$/g, "");
  }
}

function enviarJSON(res, status, datos) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
  res.end(JSON.stringify(datos));
}

function apiRecomendaciones(req, res) {
  if (req.method !== "POST") return enviarJSON(res, 405, { error: "Usa POST." });
  if (!permitir(req.socket.remoteAddress || "local")) {
    return enviarJSON(res, 429, { error: "Demasiadas peticiones. Prueba en un minuto." });
  }
  let cuerpo = "";
  let demasiado = false;
  req.on("data", (trozo) => {
    cuerpo += trozo;
    if (cuerpo.length > MAX_CUERPO) {
      demasiado = true;
      req.destroy();
    }
  });
  req.on("end", async () => {
    if (demasiado) return;
    try {
      const agregados = validarAgregados(JSON.parse(cuerpo || "{}"));
      const recomendaciones = await pedirRecomendaciones(agregados);
      enviarJSON(res, 200, { recomendaciones });
    } catch (err) {
      if (err instanceof SyntaxError) return enviarJSON(res, 400, { error: "El cuerpo no es JSON válido." });
      const { status, mensaje } = respuestaDeError(err);
      console.warn("[recomendaciones]", status, err && err.message);
      enviarJSON(res, status, { error: mensaje });
    }
  });
}

let options;
try {
  options = {
    key: fs.readFileSync(path.join(CERT_DIR, "localhost.key")),
    cert: fs.readFileSync(path.join(CERT_DIR, "localhost.crt")),
  };
} catch {
  console.error(
    "No encuentro los certificados en " + CERT_DIR + ".\n" +
      "Instálalos una vez con:  npx office-addin-dev-certs install"
  );
  process.exit(1);
}

https
  .createServer(options, (req, res) => {
    const urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
    if (urlPath === "/api/recomendaciones") return apiRecomendaciones(req, res);
    const file = path.normalize(path.join(ROOT, urlPath === "/" ? "taskpane.html" : urlPath));
    if (!file.startsWith(ROOT + path.sep)) {
      res.writeHead(403).end();
      return;
    }
    fs.readFile(file, (err, data) => {
      if (err) {
        res.writeHead(404).end("No encontrado");
        return;
      }
      res.writeHead(200, {
        "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
        "Cache-Control": "no-cache",
      });
      res.end(data);
    });
  })
  .listen(PORT, "127.0.0.1", () => {
    console.log(`Análisis de reservas servido en https://localhost:${PORT}`);
    if (!process.env.ANTHROPIC_API_KEY) console.log("Aviso: falta ANTHROPIC_API_KEY en .env; las recomendaciones no funcionarán.");
  });
