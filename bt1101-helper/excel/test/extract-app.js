// Writes the website's main script to excel/.build/app-src.js (the line numbers OPS.md uses).
const fs = require("fs");
const path = require("path");
const { extractWebsite } = require("../../pdf/src/extract.js");

const site = extractWebsite(path.join(__dirname, "..", "..", "index.html"));
fs.mkdirSync(path.join(__dirname, "..", ".build"), { recursive: true });
fs.writeFileSync(path.join(__dirname, "..", ".build", "app-src.js"), site.scripts[site.scripts.length - 1]);
console.log("Wrote excel/.build/app-src.js");
