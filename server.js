// Servidor HTTPS local mínimo (sin dependencias) para servir el complemento en https://localhost:3000
// Usa los certificados de desarrollo de Office: npx office-addin-dev-certs install
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = Number(process.env.PORT) || 3000;
const ROOT = path.join(__dirname, "src");
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".png": "image/png",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
};

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
  .listen(PORT, "127.0.0.1", () => console.log(`Respuestas IA servido en https://localhost:${PORT}`));
