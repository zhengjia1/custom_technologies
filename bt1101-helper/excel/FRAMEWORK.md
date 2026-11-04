# Framework API (VBA)

The website's script manipulates a DOM. The Excel edition keeps that model: the
website's own `<body>` markup is parsed at startup into a small VBA DOM (a port of
`pdf/src/minidom.js`), operation code is ported nearly line-by-line against it, and a
renderer draws the visible DOM into the app window (MSForms) using the website's CSS.
Operation modules therefore never touch MSForms.

## DOM (`modDom.bas`, `clsNode.cls`, `clsEvent.cls`)

One document per workbook session: `modDom.Doc` (a `clsNode` with `NodeType = 9`) and a
window-like event target `modDom.Win` (the parent of `Doc` on event paths, like `window`).

### clsNode — elements (1), text (3), comments (8), document (9), fragments (11)

JS → VBA naming is mechanical: `el.classList.add("x")` → `el.AddClass "x"`,
`el.textContent = s` → `el.TextContent = s`, `$("#id")` → `Q("#id")`.

Properties
- `NodeType`, `LocalName` (lowercase), `TagName` (uppercase), `Uid` (unique Long per node)
- `Id`, `ClassName`, `Title`, `Placeholder`, `TypeName` (input `type`), `HtmlFor` (Get/Let)
- `Data` (text/comment content), `TextContent` (Get/Let), `InnerHtml` (Get/Let), `OuterHtml`
- `Value` (Get/Let) with minidom semantics for input/textarea/select/option/button;
  `DefaultValue`; `Checked`; `SelectedIndex`; `Options` (clsList of option nodes);
  `Disabled`, `ReadOnly`, `Hidden`, `IsOpen` (`<details>` open; `Open` is reserved in VBA) (Get/Let, reflect attributes)
- `ParentNode`, `ParentElement`, `ChildNodes` / `Children` (clsList snapshots),
  `FirstChild`, `LastChild`, `FirstElementChild`, `LastElementChild`, `NextSibling`,
  `PreviousSibling`, `NextElementSibling`, `PreviousElementSibling`, `IsConnected`
- `Version` — increments whenever this node's attributes, text, value or children change

Methods
- Attributes: `GetAttribute(name)` (Variant: `Null` when absent), `HasAttribute`,
  `SetAttribute`, `RemoveAttribute`, `ToggleAttribute(name, [force])`, `AttributeNames`
- Dataset: `DataGet(camelKey)` (Variant `Null` when absent), `DataSet(camelKey, value)`,
  `DataRemove(camelKey)` (`dataset.answerVariantId` ↔ `data-answer-variant-id`)
- Classes: `HasClass`, `AddClass(ParamArray)`, `RemoveClass(ParamArray)`,
  `ToggleClass(name, [force]) As Boolean`, `ReplaceClass(old, new)`
- Inline style: `SetStyle(name, value)`, `GetStyle(name)`, `RemoveStyle(name)` (custom
  properties such as `--sidebar-width` included)
- Tree: `AppendChild`, `InsertBefore(node, ref)`, `RemoveChild`, `ReplaceChild`, `Remove`,
  `Append(ParamArray nodesOrStrings)`, `Prepend`, `Before`, `After`, `ReplaceWith`,
  `ReplaceChildren(ParamArray)`, `InsertAdjacentHtml(position, html)`,
  `InsertAdjacentElement(position, node)`, `CloneNode([deep])`, `Contains(other)`
- Queries: `QuerySelector(sel)` (Nothing when none), `QuerySelectorAll(sel)` (clsList),
  `Matches(sel)`, `Closest(sel)`, `GetElementById` (document)
- Behaviour: `Click` (dispatches click; checkbox/radio toggling with input+change events,
  `<summary>` toggles its `<details>`, `<label>` forwards to its control), `Focus`, `Blur`,
  `DispatchEvent(evt) As Boolean` (False when default prevented)
- Listeners: `AddEventListener(type, handler, [capture], [once])`,
  `RemoveEventListener(type, handler, [capture])`. `handler` is a procedure name such as
  `"opShapiro.OnInput"`, called as `Application.Run handler, evt`.

### modDom

- `Doc`, `Win`, `Q(selector)` (= `document.querySelector`), `QA(selector)` (clsList),
  `ById(id)`, `CreateElement(tag)`, `CreateTextNode(text)`, `CreateFragment()`,
  `ParseFragment(html) As clsNode`, `LoadDocument(bodyHtml)`, `ActiveElement`
- Events: `NewEvent(type, [bubbles], [cancelable]) As clsEvent`,
  `NewKeyEvent(type, key, [shift], [ctrl], [alt])`, `NewMouseEvent(type, [clientX])`
- Timers (virtual clock; see "Settling"): `SetTimeout(handler, delayMs, [arg]) As Long`,
  `ClearTimeout(id)`, `RequestAnimationFrame(handler)`, `QueueMicrotask(handler, [arg])`,
  `Settle([horizonMs])`, `NextTimerDelay() As Long` (-1 when none), `AdvanceClock()`.
  Handlers are procedure names called with `Application.Run handler[, arg]`.
- Mutation hook: `SetMutationHook(handler)` — called as `Application.Run handler, parent,
  addedNodes` after nodes are inserted into the connected document (replaces the
  website's MutationObserver uses).
- `Version` — increments on any mutation anywhere (the renderer redraws when it changes).
- `EscapeHtml(text)` — the website's `escapeHtml`.

### clsEvent

`EventType`, `Target`, `CurrentTarget`, `Bubbles`, `Cancelable`, `DefaultPrevented`,
`Key`, `ShiftKey`, `CtrlKey`, `AltKey`, `MetaKey`, `IsComposing`, `ClientX`, `ClientY`,
`Detail` (Variant), `EventPhase`; `PreventDefault`, `StopPropagation`,
`StopImmediatePropagation`.

### Implementation notes (as built)

- Also available: `modDom.Body`, `Head`, `DocumentElement`, `CreateComment`, `ResetDocument`,
  `CancelAnimationFrame`, `ClockNow`, `CssEscape`, `DecodeEntities`; clsNode `NodeName`,
  `NodeValue`, `InnerText`, `Name`, `Href`, `Role`, `TabIndex`, `Multiple`, `Required`,
  `DefaultChecked`, `DefaultSelected`, `Selected`/`Text`/`Label`/`Index` (options),
  `SelectedOption`, `LabelControl`, `ChildElementCount`, `HasChildNodes`,
  `GetElementsByTagName`, `InsertAdjacentText`, `DataHas`, `StyleNames`, `StyleText`,
  `HasListeners`, `OuterHtml` (Let); clsEvent `Code`, `Button`, `InputType`.
- `NewEvent(type)` defaults to **not** bubbling/cancelable (like `new Event(type)`); pass
  `True` for `{ bubbles: true }`. `NewKeyEvent`/`NewMouseEvent` bubble and are cancelable.
- `Focus`/`Blur` only move the active element (no events), as in minidom; the renderer
  dispatches focus/blur events for real user focus changes.
- The mutation hook is delivered as a microtask (on the next `Settle`), once per insertion
  into the connected document, not during `LoadDocument`.
- `RequestAnimationFrame` handlers take no argument and run 16 ms later on the virtual
  clock; `Settle h` runs everything due by clock-at-start + h.
- **Errors in handlers:** a handler invoked through `Application.Run` that raises an
  unhandled error shows VBA's error dialog; the dispatcher cannot catch it. Every Public
  event handler and timer/microtask procedure must trap its own errors:
  `On Error GoTo Failed` ... `Exit Sub` / `Failed: ReportError "Module.Proc"` (modShared).
- Internal helpers named `Dom*`, `SbAdd` and `Friend` members of clsNode are private to
  the DOM.

### Settling

Like the website in a browser, an event handler may schedule work (`setTimeout(fn, 0)`,
`requestAnimationFrame`). After every user action the framework calls `Settle 700`:
microtasks, then timers due within 700 ms of the virtual clock, in due order. Longer
timers (toasts, "Copied" captions) run when Excel's `Application.OnTime` fires.
Async/await code in the website becomes ordinary synchronous VBA.
