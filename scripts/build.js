/*
 * Maakt een installatiepakket voor hosting op een eigen server.
 *
 *   node scripts/build.js https://intranet.kantoor.nl/addins/definities/
 *
 * Resultaat in dist/:
 *   dist/web/          kopiëren naar de webserver (op het opgegeven adres)
 *   dist/manifest.xml  uploaden in het Microsoft 365-beheercentrum
 */
"use strict";

const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const dist = path.join(root, "dist");

function fail(msg) {
  console.error(msg);
  console.error("Gebruik: node scripts/build.js https://server.kantoor.nl/pad/naar/definities/");
  process.exit(1);
}

let base = process.argv[2];
if (!base) fail("Geef het HTTPS-adres op waar de webbestanden komen te staan.");
let url;
try {
  url = new URL(base);
} catch (e) {
  fail("Ongeldig adres: " + base);
}
if (url.protocol !== "https:") fail("Het adres moet met https:// beginnen; Office laadt add-ins alleen via HTTPS.");
if (/(^|\.)example\.(com|org|net)$/.test(url.hostname)) fail("Vul het echte adres van de eigen server in, niet de plaatshouder.");
if (!base.endsWith("/")) base += "/";

const manifest = fs.readFileSync(path.join(root, "manifest.xml"), "utf8");
const current = /<SourceLocation DefaultValue="(https:\/\/[^"]*\/)taskpane\.html"/.exec(manifest);
if (!current) fail("Kon het huidige adres niet vinden in manifest.xml.");

fs.rmSync(dist, { recursive: true, force: true });
fs.cpSync(path.join(root, "web"), path.join(dist, "web"), { recursive: true });
const built = manifest.split(current[1]).join(base);
if (built.indexOf("example.com") !== -1) fail("In manifest.xml staat nog een plaatshouder-adres dat niet vervangen kon worden.");
fs.writeFileSync(path.join(dist, "manifest.xml"), built);

console.log("Klaar.");
console.log("  1. Kopieer de inhoud van dist/web/ naar " + base);
console.log("     Controle: " + base + "taskpane.html moet in de browser de demo tonen.");
console.log("  2. Upload dist/manifest.xml in het Microsoft 365-beheercentrum.");
console.log("Zie docs/installatie-it.md voor de volledige stappen.");
