/* Marketplace / hosting uchun yig'ish.
 *   node scripts/build.js --url https://addin.example.com --email support@example.com
 * (yoki ADDIN_URL va SUPPORT_EMAIL muhit o'zgaruvchilari)
 * Natija:
 *   dist/site/      — hostingga yuklanadigan fayllar
 *   dist/manifest.xml — Partner Center'ga yuklanadigan manifest
 */
"use strict";
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const SRC = path.join(ROOT, "web");
const OUT = path.join(ROOT, "dist");
const DEV_URL = "https://localhost:3000";
const TEXT = new Set([".html", ".js", ".css", ".xml"]);

function arg(name) {
  const i = process.argv.indexOf("--" + name);
  return i > 0 ? process.argv[i + 1] : undefined;
}

function fail(msg) {
  console.error("Xatolik: " + msg);
  console.error("Foydalanish: node scripts/build.js --url https://addin.example.com --email support@example.com");
  process.exit(1);
}

const url = (arg("url") || process.env.ADDIN_URL || "").replace(/\/+$/, "");
const email = arg("email") || process.env.SUPPORT_EMAIL || "";

let parsed;
try { parsed = new URL(url); } catch (e) { fail("--url noto'g'ri yoki ko'rsatilmagan"); }
if (parsed.protocol !== "https:") fail("--url https:// bilan boshlanishi kerak");
if (/^(localhost|127\.|0\.0\.0\.0)/.test(parsed.hostname)) fail("Marketplace localhost manzilini qabul qilmaydi");
if (parsed.search || parsed.hash) fail("--url da ? yoki # bo'lmasligi kerak");
if (!/^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$/.test(email)) fail("--email noto'g'ri yoki ko'rsatilmagan");

function substitute(text) {
  return text.split(DEV_URL).join(url).split("{{SUPPORT_EMAIL}}").join(email);
}

function copyDir(from, to) {
  fs.mkdirSync(to, { recursive: true });
  for (const name of fs.readdirSync(from)) {
    if (name.startsWith(".")) continue;
    const src = path.join(from, name), dst = path.join(to, name);
    if (fs.statSync(src).isDirectory()) { copyDir(src, dst); continue; }
    if (TEXT.has(path.extname(name))) fs.writeFileSync(dst, substitute(fs.readFileSync(src, "utf8")));
    else fs.copyFileSync(src, dst);
  }
}

fs.rmSync(OUT, { recursive: true, force: true });
copyDir(SRC, path.join(OUT, "site"));
// Dasturchi uchun izohlar Marketplace manifestiga kerak emas
const manifest = fs.readFileSync(path.join(ROOT, "manifest.xml"), "utf8").replace(/<!--[\s\S]*?-->\s*/g, "")
  .split("<AppDomain>" + DEV_URL + "</AppDomain>").join("<AppDomain>" + parsed.origin + "</AppDomain>");   // AppDomain — faqat origin
fs.writeFileSync(path.join(OUT, "manifest.xml"), substitute(manifest));
// Saytda ham turadi: foydalanuvchi manifestni to'g'ridan-to'g'ri havoladan yuklab oladi
fs.copyFileSync(path.join(OUT, "manifest.xml"), path.join(OUT, "site", "manifest.xml"));

// Hech qanday localhost yoki to'ldirilmagan joy qolmaganini tekshirish
const leftovers = [];
(function scan(dir) {
  for (const name of fs.readdirSync(dir)) {
    const p = path.join(dir, name);
    if (fs.statSync(p).isDirectory()) { scan(p); continue; }
    if (!TEXT.has(path.extname(name))) continue;
    const s = fs.readFileSync(p, "utf8");
    if (s.includes("localhost") || /\{\{[A-Z_]+\}\}/.test(s)) leftovers.push(path.relative(ROOT, p));
  }
})(OUT);
if (leftovers.length) fail("almashtirilmagan joylar qoldi: " + leftovers.join(", "));

console.log("Tayyor:");
console.log("  dist/site/      → " + url + "/ ga yuklang");
console.log("  dist/manifest.xml → Partner Center'ga yuklang");
console.log("Tekshirish: npx office-addin-manifest validate dist/manifest.xml -p");
