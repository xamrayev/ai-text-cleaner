/* web/ papkasini https://localhost:3000 da tarqatadi (Office ishonchli dev sertifikati bilan). */
"use strict";
const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const ROOT = path.join(__dirname, "..", "web");
const PORT = 3000;
const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml" };

devCerts.getHttpsServerOptions().then((opts) => {
  https.createServer(opts, (req, res) => {
    const url = decodeURIComponent(req.url.split("?")[0]);
    const file = path.normalize(path.join(ROOT, url === "/" ? "taskpane.html" : url));
    if (!file.startsWith(ROOT)) { res.writeHead(403); return res.end(); }
    fs.readFile(file, (err, data) => {
      if (err) { res.writeHead(404); return res.end("Topilmadi"); }
      res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream", "Cache-Control": "no-cache" });
      res.end(data);
    });
  }).listen(PORT, () => console.log(`Server: https://localhost:${PORT}/taskpane.html`));
});
