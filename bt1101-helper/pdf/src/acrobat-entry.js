// Runs inside Adobe Acrobat as a document-level script. BT_tasks is created by the
// prelude before any polyfill loads, so promises and timers share one queue.

import { createBrowser, installNumberLocale } from "./env.js";
import { createBridge } from "./bridge.js";
import { runApp } from "../.build/app.js";
import { BODY, REGISTRY, VERSION } from "../.build/site-data.js";

export function boot(acrobat) {
  const started = Date.now();
  const doc = acrobat.doc;
  const log = message => { try { acrobat.console.println(message); } catch (error) { /* console closed */ } };
  const bootErrors = [];
  installNumberLocale();
  const browser = createBrowser(BODY, BT_tasks, { reportError: error => { bootErrors.push(error); log("BT1101 error: " + (error && error.stack ? error.stack : error)); } });
  const app = runApp(browser.globals);
  browser.drain();
  const host = {
    getField: name => doc.getField(name),
    setPage: page => { doc.pageNum = page; },
    setDelay: on => { try { doc.delay = on; } catch (error) { /* older viewers */ } },
    popUp: items => acrobat.app.popUpMenuEx.apply(acrobat.app, items),
    alert: options => acrobat.app.alert(options),
    display: acrobat.display,
    log
  };
  const readyMessage = bootErrors.length
    ? "Started with an error: " + (bootErrors[0] && bootErrors[0].message ? bootErrors[0].message : bootErrors[0])
    : "Ready · " + VERSION + " · Type or paste, then press Tab to update.";
  const bridge = createBridge({ host, document: browser.document, app, browser, registry: REGISTRY, readyMessage });
  log("BT1101 timing: start-up " + (Date.now() - started) + " ms (program loaded in " + (started - BT_loadStarted) + " ms before that)");
  return bridge;
}
