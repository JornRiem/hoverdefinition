/*
 * Kleine lokale webserver om de demo te bekijken, zonder extra software.
 *
 *   npm start            → http://localhost:8080/taskpane.html
 *   npm start -- 3000    → andere poort
 *
 * Alleen bedoeld om lokaal te kijken; Word zelf laadt add-ins alleen via HTTPS.
 */
"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..", "web");
const port = Number(process.argv[2]) || 8080;
const types = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".txt": "text/plain; charset=utf-8"
};

http
  .createServer((req, res) => {
    const urlPath = decodeURIComponent(new URL(req.url, "http://localhost").pathname);
    const file = path.join(root, urlPath === "/" ? "taskpane.html" : urlPath);
    if (!file.startsWith(root + path.sep)) {
      res.writeHead(403).end();
      return;
    }
    fs.readFile(file, (err, data) => {
      if (err) {
        res.writeHead(404).end("Niet gevonden");
        return;
      }
      res.writeHead(200, { "Content-Type": types[path.extname(file)] || "application/octet-stream" });
      res.end(data);
    });
  })
  .listen(port, () => console.log("Demo: http://localhost:" + port + "/taskpane.html"));
