// Loads the website logic on minidom inside Node, without transpiling, for fast checks.

const path = require("path");
const fs = require("fs");
const esbuild = require("esbuild");
const { extractWebsite, writeAppModule } = require("../src/extract.js");

const ROOT = path.join(__dirname, "..");
const BUILD = path.join(ROOT, ".build");

function bundleForNode() {
  const site = extractWebsite(path.join(ROOT, "..", "index.html"));
  writeAppModule(site, path.join(BUILD, "app.js"));
  fs.writeFileSync(path.join(BUILD, "node-entry.js"), `
    import * as dom from "../src/minidom.js";
    import { createBrowser, installNumberLocale } from "../src/env.js";
    import { runApp } from "./app.js";
    export { dom, createBrowser, installNumberLocale, runApp };
  `);
  const out = path.join(BUILD, "node-bundle.js");
  esbuild.buildSync({ entryPoints: [path.join(BUILD, "node-entry.js")], bundle: true, format: "cjs", platform: "node", outfile: out, logLevel: "error" });
  delete require.cache[out];
  return { site, lib: require(out) };
}

function bootNodeApp() {
  const { site, lib } = bundleForNode();
  const tasks = { micro: [], timers: [], seq: 0, now: 0 };
  const errors = [];
  const browser = lib.createBrowser(site.body, tasks, { reportError: error => errors.push(error) });
  let app;
  try { app = lib.runApp(browser.globals); } catch (error) { errors.push(error); }
  browser.drain();
  return { site, lib, browser, app, errors, tasks };
}

module.exports = { bundleForNode, bootNodeApp, ROOT, BUILD };
