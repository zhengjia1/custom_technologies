// Writes the website's <body> markup to excel/.build/body.html for VBA DOM tests.
const fs = require("fs");
const path = require("path");
const { extractWebsite } = require("../../pdf/src/extract.js");

const site = extractWebsite(path.join(__dirname, "..", "..", "index.html"));
fs.mkdirSync(path.join(__dirname, "..", ".build"), { recursive: true });
fs.writeFileSync(path.join(__dirname, "..", ".build", "body.html"), site.body);
console.log("Wrote excel/.build/body.html");
