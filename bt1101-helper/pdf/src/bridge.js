// Connects Acrobat form fields to the website running on minidom. Field actions call
// commit/toggle/press/pageOpen; each turns into the DOM events a browser would fire,
// lets the website react, then copies the resulting page state back into the fields.

import { resolveRef, isRendered, visibleText, buttonCaption } from "./refs.js";

const EXPORT_PREFIX = "@";
// Website colour tokens: --c-e9f6ed (copied), --c-fce8e6 (error), --c-d3e3fd (active),
// --c-f1f3f4 (disabled), --c-1f1f1f (ink), --c-5f6368 (muted), --c-b3261e (error text).
const COLOURS = {
  done: ["RGB", 0.914, 0.965, 0.929], errorFill: ["RGB", 0.988, 0.91, 0.902], active: ["RGB", 0.827, 0.89, 0.992],
  disabled: ["RGB", 0.945, 0.953, 0.957], ink: ["RGB", 0.122, 0.122, 0.122], muted: ["RGB", 0.373, 0.388, 0.408],
  error: ["RGB", 0.702, 0.149, 0.118], transparent: ["T"]
};

export function createBridge(options) {
  const host = options.host, document = options.document, app = options.app, browser = options.browser, registry = options.registry;
  // The theme's colours, captured from the website; COLOURS holds the light-theme defaults.
  const palette = registry.palette || {};
  const tone = (name, fallback) => (palette[name] ? ["RGB"].concat(palette[name]) : fallback);
  const theme = {
    done: tone("done", COLOURS.done), errorFill: tone("errorFill", COLOURS.errorFill), active: tone("active", COLOURS.active), disabled: tone("disabled", COLOURS.disabled),
    ink: tone("ink", COLOURS.ink), muted: tone("muted", COLOURS.muted), error: tone("error", COLOURS.error), badge: tone("badge", ["RGB", 0.184, 0.49, 0.322]), badgeText: tone("badgeText", ["RGB", 1, 1, 1]),
    transparent: COLOURS.transparent
  };
  const Event = browser.globals.Event;
  const cache = {}, applied = {};
  const errors = [];
  let currentPage = -1, busy = false, toastMessage = "", toastChanged = false;

  // Buttons start with the appearance drawn at build time. Acrobat redraws a button when
  // its caption or colours change, so only real changes are sent.
  Object.keys(registry.fields).forEach(name => {
    const spec = registry.fields[name];
    if (spec.k === "button" && spec.cap !== undefined) applied[name] = { caption: spec.cap, state: "normal" };
  });

  const toast = document.getElementById("toast");
  if (toast) new browser.globals.MutationObserver(() => { toastChanged = true; }).observe(toast, { childList: true, subtree: true });

  function recordError(error) {
    const message = error && error.message ? error.message : String(error);
    errors.push(message);
    if (errors.length > 5) errors.shift();
    host.log("BT1101 error: " + message + (error && error.stack ? "\n" + error.stack : ""));
  }

  function element(ref) {
    const known = cache[ref];
    if (known && known.isConnected) return known;
    const found = resolveRef(document, ref);
    cache[ref] = found;
    return found;
  }

  function fire(target, type) { target.dispatchEvent(new Event(type, { bubbles: true })); }

  function settle() {
    browser.drain();
    for (let guard = 0; guard < 6; guard++) {
      const alerts = document.querySelectorAll(".variable-rename-alert:not(.hidden)");
      if (!alerts.length) break;
      showAlert(alerts[alerts.length - 1]);
      browser.drain();
    }
  }

  // The website's warning modals become Acrobat alerts; two-button dialogs become Yes/No.
  function showAlert(dialog) {
    const body = dialog.querySelector(".variable-rename-dialog") || dialog;
    const kicker = body.querySelector(".variable-rename-kicker");
    const title = kicker ? kicker.textContent.trim() : "BT1101 Prose Lab";
    const message = visibleText(body).replace(title, "").trim();
    const buttons = dialog.querySelectorAll("button");
    if (buttons.length <= 1) {
      host.alert({ cTitle: title, cMsg: message, nIcon: 1, nType: 0 });
      if (buttons[0]) buttons[0].click();
      return;
    }
    const decline = buttons[0], accept = buttons[buttons.length - 1];
    const answer = host.alert({ cTitle: title, cMsg: message + "\n\nYes: " + buttonCaption(accept) + "\nNo: " + buttonCaption(decline), nIcon: 2, nType: 2 });
    (answer === 4 ? accept : decline).click();
  }

  // The page for an operation in its current layout: the smallest captured layout that
  // includes every element now showing; the catch-all page (no sig) includes them all.
  function pageOfOperation(id) {
    const tracked = registry.layout && registry.layout[id];
    let best = -1, bestSize = Infinity;
    const showing = tracked ? tracked.filter(ref => { const target = element(ref); return target && isRendered(target, true); }) : [];
    for (let i = 0; i < registry.pages.length; i++) {
      const page = registry.pages[i];
      if (page.op !== id) continue;
      if (!tracked) return i;
      const size = page.sig ? page.sig.length : Infinity;
      if (page.sig && showing.some(ref => page.sig.indexOf(ref) < 0)) continue;
      if (best < 0 || size < bestSize) { best = i; bestSize = size; }
    }
    return best;
  }

  function openOperation(id) {
    const entry = id === "universal" ? document.querySelector(".universal-entry") : document.querySelector('#tutorialList .operation[data-id="' + id + '"]');
    if (entry) entry.click();
  }

  // One console line per action: time in the website's code, in updating the PDF's
  // fields (and how many Acrobat calls that took), and in Acrobat's redraw.
  let fieldCalls = 0;
  function timed(label, work) {
    const started = Date.now();
    fieldCalls = 0;
    const result = work();
    host.log("BT1101 timing: " + label + " " + (Date.now() - started) + " ms" + (result ? " (" + result + ")" : ""));
  }

  function run(action, context, label) {
    if (busy) return;
    busy = true;
    const started = Date.now();
    fieldCalls = 0;
    const writes = browser.clipboard.writes.length;
    const before = app.state.id;
    toastChanged = false;
    try { action(); settle(); } catch (error) { recordError(error); }
    const website = Date.now();
    busy = false;
    // Acrobat redraws each field as it changes; hold redrawing until all fields are updated.
    host.setDelay(true);
    if (toastChanged && toast) toastMessage = toast.textContent.trim();
    const after = app.state.id;
    const target = pageOfOperation(after);
    // A different operation, or the same one in a different layout, lives on another page.
    if (target >= 0 && target !== currentPage && (after !== before || (registry.pages[currentPage] && registry.pages[currentPage].op === after))) {
      currentPage = target;
      host.setPage(target);
    }
    let synced = website;
    try { syncPage(currentPage, context); synced = Date.now(); } finally { host.setDelay(false); }
    if (browser.clipboard.writes.length > writes) deliverClipboard(browser.clipboard.writes[browser.clipboard.writes.length - 1], context);
    const done = Date.now();
    host.log("BT1101 timing: " + (label || (context && context.name) || "action") + " " + (done - started) + " ms (website " + (website - started) + " ms, fields " + (synced - website) + " ms with " + fieldCalls + " Acrobat calls, redraw " + (done - synced) + " ms)");
  }

  // Dropdowns are a value box plus a button that opens Acrobat's pop-up menu: the current
  // choice is ticked, unavailable options are greyed out, plot types sit under headings.
  function menuItems(spec) {
    if (spec.k === "plot") {
      const select = element("#plotWorkflow"), search = element("#plotWorkflowSearch");
      const current = search ? search.value.trim() : "";
      const items = [];
      let category = null;
      select._options().forEach((option, index) => {
        const next = option.dataset.category || "";
        if (next !== category) { if (category !== null) items.push({ cName: "-" }); category = next; items.push({ cName: next, bEnabled: false }); }
        const label = option.textContent.trim();
        items.push({ cName: "   " + label, cReturn: EXPORT_PREFIX + index, bMarked: label === current });
      });
      return items;
    }
    const target = element(spec.r);
    if (!target) return [];
    const selected = target.selectedIndex;
    return target._options().map((option, index) => ({ cName: option.textContent.replace(/\s+/g, " ").trim() || " ", cReturn: EXPORT_PREFIX + index, bMarked: index === selected, bEnabled: !option.disabled }));
  }

  // Acrobat cannot write to the clipboard, so the exact copied text goes into the copy box
  // beside the button and is focused, ready for Ctrl+C.
  function deliverClipboard(text, context) {
    let clip = context && context.spec && context.spec.c;
    if (!clip && currentPage >= 0) clip = registry.pages[currentPage].clip;
    if (!clip) return;
    const field = fieldObject(clip);
    if (!field) return;
    field.value = text;
    setDisplay(clip, field, true);
    toastMessage = "Copied text is in the yellow box: click it, press Ctrl+A, then Ctrl+C.";
    const status = registry.pages[currentPage] && registry.pages[currentPage].fields.filter(name => registry.fields[name].k === "status")[0];
    if (status) syncField(status, null);
    try { field.setFocus(); } catch (error) { recordError(error); }
  }

  function setDisplay(name, field, visible) {
    const state = applied[name] || (applied[name] = {});
    if (state.visible === visible) return;
    state.visible = visible;
    fieldCalls++;
    field.display = visible ? host.display.visible : host.display.hidden;
  }
  function setProperty(name, field, key, value, apply) {
    const state = applied[name] || (applied[name] = {});
    const signature = typeof value === "object" ? JSON.stringify(value) : value;
    if (state[key] === signature) return;
    state[key] = signature;
    fieldCalls++;
    apply(field, value);
  }
  function setValue(name, field, value, context) {
    if (context && context.name === name && context.event) {
      context.event.value = value;
      (applied[name] || (applied[name] = {})).value = value;
      return;
    }
    setProperty(name, field, "value", value, (target, text) => { if (String(target.value) !== String(text)) target.value = text; });
  }

  function optionLabel(option) {
    const label = option.textContent.replace(/\s+/g, " ").trim();
    return option.disabled ? label + " (not applicable)" : label;
  }

  function stateColour(target, fallback) {
    if (target.closest(".is-copy-complete") || target.closest(".lp-collapsible.is-complete")) return theme.done;
    if (target.closest(".is-error") || target.querySelector(".is-error")) return theme.errorFill;
    return fallback;
  }

  // Acrobat's getField is a costly call; each field object is fetched once.
  const fieldObjects = {};
  function fieldObject(name) {
    if (!(name in fieldObjects)) { fieldObjects[name] = host.getField(name); fieldCalls++; }
    return fieldObjects[name];
  }

  function syncField(name, context) {
    const spec = registry.fields[name];
    const field = fieldObject(name);
    if (!spec || !field) return;
    if (spec.k === "clip") return;
    if (spec.k === "status") {
      const text = errors.length ? "Error: " + errors[errors.length - 1] : toastMessage || options.readyMessage || "";
      setValue(name, field, text, null);
      setProperty(name, field, "text", errors.length ? "error" : "ok", (target, kind) => { target.textColor = kind === "error" ? theme.error : theme.muted; });
      return;
    }
    const target = element(spec.k === "plot" ? "#plotWorkflowSearch" : spec.r);
    const visible = Boolean(target) && isRendered(target, true);
    setDisplay(name, field, spec.k === "badge" ? visible && target.classList.contains("is-complete") : visible);
    if (!target || !visible) return;
    const own = spec.f ? ["RGB"].concat(spec.f) : theme.transparent;
    if (spec.k === "text" || spec.k === "search") {
      setValue(name, field, target.value, context);
      // Disabled inputs are grey; the website's invalid constraint cells are red.
      const look = target.disabled || target.readOnly ? "disabled" : target.classList.contains("input-error") || target.getAttribute("aria-invalid") === "true" ? "invalid" : "normal";
      setProperty(name, field, "look", look, (f, value) => { f.readonly = value === "disabled"; f.fillColor = value === "disabled" ? theme.disabled : value === "invalid" ? theme.errorFill : own; });
    } else if (spec.k === "select") {
      const option = target._options()[target.selectedIndex];
      setValue(name, field, option ? option.textContent.replace(/\s+/g, " ").trim() : "", null);
      setProperty(name, field, "look", target.disabled ? "disabled" : "normal", (f, value) => { f.textColor = value === "disabled" ? theme.muted : spec.t ? ["RGB"].concat(spec.t) : theme.ink; f.fillColor = value === "disabled" ? theme.disabled : own; });
    } else if (spec.k === "plot") {
      setValue(name, field, target.value.trim(), null);
    } else if (spec.k === "pick") {
      setProperty(name, field, "readonly", spec.plot ? false : Boolean(target.disabled), (f, value) => { f.readonly = value; });
    } else if (spec.k === "check") {
      setValue(name, field, target.checked ? "Yes" : "Off", null);
      setProperty(name, field, "readonly", Boolean(target.disabled), (f, value) => { f.readonly = value; });
    } else if (spec.k === "button") {
      // Keep the website's own caption (and the drawn icon) until the website changes it.
      const raw = target.textContent.replace(/\s+/g, " ").trim();
      setProperty(name, field, "caption", spec.cap !== undefined && raw === spec.cap ? spec.cap : buttonCaption(target), (f, caption) => { f.buttonSetCaption(caption); });
      const disabled = Boolean(target.disabled);
      const active = target.classList.contains("active") || target.getAttribute("aria-pressed") === "true";
      const done = /^Copied$/i.test(target.textContent.trim()) || Boolean(target.closest(".is-copy-complete"));
      setProperty(name, field, "state", disabled ? "disabled" : active ? "active" : done ? "done" : "normal", (f, state) => {
        f.readonly = state === "disabled";
        f.textColor = state === "disabled" ? theme.muted : spec.t ? ["RGB"].concat(spec.t) : theme.ink;
        f.fillColor = state === "active" ? theme.active : state === "done" ? theme.done : own;
      });
    } else if (spec.k === "badge") {
      // --c-2f7d52, the website's completion-badge green.
      setValue(name, field, "Done!", null);
      setProperty(name, field, "fill", "badge", (f) => { f.fillColor = theme.badge; f.textColor = theme.badgeText; });
    } else if (spec.k === "label") {
      return;
    } else if (spec.k === "out" || spec.k === "note" || spec.k === "cardbody") {
      let text;
      if (spec.k === "note") text = target.textContent.replace(/\s+/g, " ").trim();
      else if (spec.k === "cardbody") {
        text = visibleText(target, ".sensitivity-prose-head");
        // More sections than the page has room for: the last box carries the rest.
        if (spec.last) for (let next = target.nextElementSibling; next; next = next.nextElementSibling) if (isRendered(next, true)) text += "\n\n" + visibleText(next);
      } else text = visibleText(target);
      setValue(name, field, text, null);
      if (spec.k !== "note") setProperty(name, field, "fill", JSON.stringify(stateColour(target, own)), (f, colour) => { f.fillColor = JSON.parse(colour); });
    }
  }

  function syncPage(page, context) {
    const entry = registry.pages[page];
    if (!entry) return;
    for (let i = 0; i < entry.fields.length; i++) {
      try { syncField(entry.fields[i], context); } catch (error) { recordError(error); }
    }
  }

  function decodeIndex(value) {
    const text = String(value);
    return text.charAt(0) === EXPORT_PREFIX && /^\d+$/.test(text.slice(1)) ? Number(text.slice(1)) : -1;
  }

  return {
    errors,
    pageOpen(page) {
      currentPage = page;
      const id = registry.pages[page] && registry.pages[page].op;
      if (id && app.state.id !== id) { run(() => openOperation(id), null, "open " + id); return; }
      const layoutPage = id ? pageOfOperation(id) : page;
      if (layoutPage >= 0 && layoutPage !== page) { currentPage = layoutPage; host.setPage(layoutPage); return; }
      timed("show " + id, () => { host.setDelay(true); try { syncPage(page, null); } finally { host.setDelay(false); } return fieldCalls + " Acrobat calls"; });
    },
    commit(name, value, event) {
      const spec = registry.fields[name];
      if (!spec) return;
      this.catchUp(name);
      // Acrobat separates the lines of a multi-line field with "\r"; a browser's textarea
      // gives the website "\n", which its parsers split on.
      if (typeof value === "string") value = value.replace(/\r\n?/g, "\n");
      const context = { name, event, spec };
      run(() => {
        const target = element(spec.r);
        if (spec.k === "text" || spec.k === "search") {
          if (!target || target.disabled || target.value === value) return;
          if (target.localName === "textarea") fire(target, "paste");
          target.value = value;
          fire(target, "input");
          fire(target, "change");
          // Searching opens the best match, as the website does when you press Enter.
          if (spec.k === "search" && value.trim()) {
            browser.drain();
            const match = document.querySelector("#rankedSearchResults .operation, #rankedSearchResults button") || document.querySelector("#tutorialList .operation:not(.hidden)");
            if (match) match.click();
          }
        } else if (spec.k === "select") {
          const options = target ? target._options() : [];
          let index = decodeIndex(value);
          if (index < 0) for (let i = 0; i < options.length; i++) if (optionLabel(options[i]) === value) index = i;
          const option = options[index];
          if (!option || option.disabled || target.disabled || target.selectedIndex === index) return;
          target.value = option.value;
          fire(target, "input");
          fire(target, "change");
        } else if (spec.k === "plot") {
          const options = element("#plotWorkflow")._options();
          const index = decodeIndex(value);
          const search = element("#plotWorkflowSearch");
          if (!options[index] || !search || search.value.trim() === options[index].textContent.trim()) return;
          search.value = options[index].textContent.trim();
          fire(search, "input");
          fire(search, "change");
          document.body.click();
        }
      }, context);
    },
    toggle(name, checked) {
      const spec = registry.fields[name];
      if (!spec) return;
      this.catchUp(name);
      run(() => { const target = element(spec.r); if (target && !target.disabled && target.checked !== checked) target.click(); }, { name, spec });
    },
    press(name) {
      const spec = registry.fields[name];
      if (!spec) return;
      this.catchUp(name);
      if (spec.k === "pick") {
        const display = registry.fields[spec.d];
        const items = menuItems(display);
        if (!items.length) return;
        const choice = host.popUp(items);
        if (choice) this.commit(spec.d, choice, null);
        return;
      }
      run(() => { const target = element(spec.r); if (target && !target.disabled && isRendered(target, true)) target.click(); }, { name, spec });
    },
    syncPage(page) { syncPage(page === undefined ? currentPage : page, null); },
    // Page opens wait until scrolling stops; a field used before then brings its page up to date.
    catchUp(name) {
      const match = /^p(\d+)\./.exec(name);
      const page = match ? Number(match[1]) : -1;
      if (page < 0 || page === currentPage) return;
      const id = registry.pages[page] && registry.pages[page].op;
      currentPage = page;
      if (id && app.state.id !== id) run(() => openOperation(id), null, "open " + id);
      else syncPage(page, null);
    },
    currentPage() { return currentPage; }
  };
}
