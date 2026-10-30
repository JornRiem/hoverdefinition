/* global Office, Word, Definitions, createDemoHost */
"use strict";

var state = {
  host: null,
  index: Definitions.buildIndex([]),
  stack: [],
  expanded: new Set(),
  savedPlace: null,
  ignoreSelectionUntil: 0,
  lastScan: 0
};

var RESCAN_AFTER_MS = 60 * 1000;

var $ = function (id) { return document.getElementById(id); };

/* ---------- Koppeling met Word ---------- */

function createWordHost() {
  var listeners = [];
  var timer = null;

  function readSelection() {
    return Word.run(async function (context) {
      var sel = context.document.getSelection();
      sel.load("text");
      var para = sel.paragraphs.getFirst();
      para.load("text");
      var before = para.getRange("Start").expandTo(sel.getRange("Start"));
      before.load("text");
      await context.sync();
      return { selected: sel.text, paragraph: para.text, offset: before.text.length };
    });
  }

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function () {
    clearTimeout(timer);
    timer = setTimeout(function () {
      readSelection()
        .then(function (info) { listeners.forEach(function (cb) { cb(info); }); })
        .catch(function () { /* bijv. selectie in een tekstvak: negeren */ });
    }, 150);
  });

  async function selectByIndex(context, def) {
    var s = def.source;
    if (s.kind === "table") {
      var tables = context.document.body.tables;
      tables.load("items/rowCount");
      await context.sync();
      var table = tables.items[s.table];
      if (!table || s.row >= table.rowCount) return false;
      var cell = table.getCell(s.row, 0);
      cell.body.load("text");
      await context.sync();
      if (Definitions.cleanTerm(cell.body.text) !== def.term) return false;
      cell.body.select();
    } else {
      var paras = context.document.body.paragraphs;
      paras.load("items/text");
      await context.sync();
      var para = paras.items[s.paragraph];
      if (!para || para.text.indexOf(def.term) === -1) return false;
      para.select();
    }
    await context.sync();
    return true;
  }

  return {
    scan: function () {
      return Word.run(async function (context) {
        var body = context.document.body;
        var tables = body.tables;
        tables.load("items/values");
        var paras = body.paragraphs;
        paras.load("items/text,items/tableNestingLevel");
        await context.sync();
        return {
          tables: tables.items.map(function (t) { return t.values; }),
          paragraphs: paras.items
            .map(function (p, i) { return { text: p.text, index: i, inTable: p.tableNestingLevel > 0 }; })
            .filter(function (p) { return !p.inTable; })
        };
      });
    },
    onSelection: function (cb) { listeners.push(cb); },
    goTo: function (def) {
      return Word.run(async function (context) {
        var ok = false;
        try { ok = await selectByIndex(context, def); } catch (e) { ok = false; }
        if (ok) return;
        // Document is gewijzigd sinds het inlezen: zoek de term zelf op.
        var hits = context.document.body.search(def.term, { matchCase: true, matchWholeWord: true });
        hits.load("items");
        await context.sync();
        if (hits.items.length) {
          hits.items[0].select();
          await context.sync();
        }
      });
    },
    savePlace: function () {
      return Word.run(async function (context) {
        var range = context.document.getSelection();
        context.trackedObjects.add(range);
        await context.sync();
        return range;
      });
    },
    returnToText: function (range) {
      if (!range) return Promise.resolve();
      return Word.run(range, async function (context) {
        range.select();
        await context.sync();
      });
    }
  };
}

/* ---------- Inlezen ---------- */

async function rescan(quiet) {
  if (!quiet) $("status").textContent = "Document wordt ingelezen…";
  try {
    var doc = await state.host.scan();
    var defs = Definitions.extractDefinitions(doc);
    state.index = Definitions.buildIndex(defs);
    state.lastScan = Date.now();
    var n = state.index.byTerm.size;
    $("status").textContent = n
      ? n + (n === 1 ? " term gevonden." : " termen gevonden.")
      : "Geen definities gevonden. De add-in zoekt naar een tabel met de term in de linkerkolom en de betekenis in de rechterkolom, en naar bepalingen als “Term” means … of “Term” betekent ….";
    renderList();
    if (state.stack.length) {
      state.stack = state.stack.filter(function (t) { return state.index.byTerm.has(t); });
      renderDetail();
    }
  } catch (e) {
    $("status").textContent = "Inlezen mislukt: " + (e && e.message ? e.message : e);
  }
}

/* ---------- Weergave ---------- */

function sourceLabel(def) {
  if (def.source.kind === "table") return "Definitietabel";
  if (def.source.kind === "paragraph") return "Definitiebepaling";
  return "Gedefinieerd in de tekst";
}

function showTerm(term, mode) {
  if (!state.index.byTerm.has(term)) return;
  if (mode === "reset") {
    state.stack = [term];
  } else {
    var at = state.stack.indexOf(term);
    state.stack = at >= 0 ? state.stack.slice(0, at + 1) : state.stack.concat(term);
  }
  state.expanded.clear();
  renderDetail();
}

// Sleutels van alle uitvouwbare termen van de getoonde term.
function allExpansionKeys(term) {
  var keys = [];
  state.index.byTerm.get(term).forEach(function (def, d) {
    keys = keys.concat(Definitions.expansionKeys(def, state.index, [term], term + "#" + d));
  });
  return keys;
}

/*
 * Toont een definitietekst met klikbare termen. Een klik vouwt de definitie
 * van die term tussen [ ] in de tekst zelf uit, zodat de hele zin in één keer
 * te lezen is.
 */
function renderText(parent, def, ancestors, key) {
  Definitions.tokenize(def.text, state.index, Definitions.ownTerms(def)).forEach(function (part, i) {
    if (!part.term) {
      parent.appendChild(document.createTextNode(part.text));
      return;
    }
    var k = key + "/" + i;
    var cycle = ancestors.indexOf(part.term) !== -1;
    var open = !cycle && state.expanded.has(k);

    var link = document.createElement("button");
    link.className = "term-link" + (open ? " open" : "") + (cycle ? " cycle" : "");
    link.textContent = part.text;
    link.setAttribute("aria-expanded", String(open));
    link.title = cycle
      ? "Kringverwijzing: " + part.term + " wordt hierboven al uitgelegd. Klik om los te openen."
      : (open ? "Inklappen" : "Uitvouwen: " + part.term);
    link.addEventListener("click", function () {
      if (cycle) return showTerm(part.term, "push");
      if (open) state.expanded.delete(k);
      else state.expanded.add(k);
      renderDetail();
    });
    parent.appendChild(link);
    if (!open) return;

    var box = document.createElement("span");
    box.className = "expansion";
    var start = document.createElement("span");
    start.className = "bracket";
    start.textContent = "[";
    box.appendChild(start);
    var jump = document.createElement("button");
    jump.className = "open-term";
    jump.textContent = "↗";
    jump.title = "Open " + part.term + " als losse term";
    jump.addEventListener("click", function () { showTerm(part.term, "push"); });
    box.appendChild(jump);
    renderText(box, state.index.byTerm.get(part.term)[0], ancestors.concat(part.term), k);
    var end = document.createElement("span");
    end.className = "bracket";
    end.textContent = "]";
    box.appendChild(end);
    parent.appendChild(box);
  });
}

function toggleExpandAll() {
  var term = state.stack[state.stack.length - 1];
  if (!term) return;
  var keys = allExpansionKeys(term);
  var allOpen = keys.every(function (k) { return state.expanded.has(k); });
  state.expanded.clear();
  if (!allOpen) keys.forEach(function (k) { state.expanded.add(k); });
  renderDetail();
}

function renderDetail() {
  var term = state.stack[state.stack.length - 1];
  $("detail").hidden = !term;
  $("hint").hidden = Boolean(term);
  if (!term) return;

  var crumbs = $("crumbs");
  crumbs.replaceChildren();
  if (state.stack.length > 1) {
    state.stack.forEach(function (t, i) {
      if (i > 0) {
        var sep = document.createElement("span");
        sep.className = "sep";
        sep.textContent = "›";
        crumbs.appendChild(sep);
      }
      if (i === state.stack.length - 1) {
        var cur = document.createElement("span");
        cur.className = "current";
        cur.textContent = t;
        crumbs.appendChild(cur);
      } else {
        var b = document.createElement("button");
        b.textContent = t;
        b.addEventListener("click", function () {
          state.stack = state.stack.slice(0, i + 1);
          state.expanded.clear();
          renderDetail();
        });
        crumbs.appendChild(b);
      }
    });
  }

  $("term-title").textContent = term;

  var keys = allExpansionKeys(term);
  var expandAll = $("expand-all");
  expandAll.hidden = !keys.length;
  expandAll.textContent = keys.length && keys.every(function (k) { return state.expanded.has(k); })
    ? "Alles inklappen"
    : "Alles uitvouwen";

  var box = $("definitions");
  box.replaceChildren();
  state.index.byTerm.get(term).forEach(function (def, d) {
    var div = document.createElement("div");
    div.className = "definition";
    var label = document.createElement("span");
    label.className = "source";
    label.textContent = sourceLabel(def);
    div.appendChild(label);
    renderText(div, def, [term], term + "#" + d);
    box.appendChild(div);
  });
}

function filteredTerms() {
  var q = $("search").value.trim().toLowerCase();
  return Array.from(state.index.byTerm.keys())
    .filter(function (t) { return !q || t.toLowerCase().indexOf(q) !== -1; })
    .sort(function (a, b) { return a.localeCompare(b, "nl"); });
}

function renderList() {
  var list = $("term-list");
  list.replaceChildren();
  filteredTerms().forEach(function (t) {
    var li = document.createElement("li");
    var b = document.createElement("button");
    b.textContent = t;
    b.addEventListener("click", function () { showTerm(t, "reset"); });
    li.appendChild(b);
    list.appendChild(li);
  });
}

/* ---------- Gebeurtenissen ---------- */

function onSelection(info) {
  if (Date.now() < state.ignoreSelectionUntil) return;
  if (Date.now() - state.lastScan > RESCAN_AFTER_MS) rescan(true);
  var term = Definitions.resolve(info.selected, info.paragraph, info.offset, state.index);
  var current = state.stack[state.stack.length - 1];
  // Staat de cursor niet op een term, dan blijft de huidige definitie staan.
  if (term && term !== current) showTerm(term, "reset");
}

async function goToDefinition() {
  var term = state.stack[state.stack.length - 1];
  if (!term) return;
  try {
    state.savedPlace = await state.host.savePlace();
    state.ignoreSelectionUntil = Date.now() + 1000;
    await state.host.goTo(state.index.byTerm.get(term)[0]);
    $("return").hidden = !state.savedPlace;
  } catch (e) {
    $("status").textContent = "Kon niet naar de definitie springen: " + (e && e.message ? e.message : e);
  }
}

async function returnToText() {
  state.ignoreSelectionUntil = Date.now() + 1000;
  try {
    await state.host.returnToText(state.savedPlace);
  } catch (e) { /* plek bestaat niet meer */ }
  state.savedPlace = null;
  $("return").hidden = true;
}

function start(host) {
  state.host = host;
  host.onSelection(onSelection);
  $("refresh").addEventListener("click", function () { rescan(false); });
  $("goto").addEventListener("click", goToDefinition);
  $("expand-all").addEventListener("click", toggleExpandAll);
  $("return").addEventListener("click", returnToText);
  $("search").addEventListener("input", renderList);
  $("search").addEventListener("keydown", function (e) {
    if (e.key !== "Enter") return;
    var first = filteredTerms()[0];
    if (first) showTerm(first, "reset");
  });
  rescan(false);
}

function startDemo() {
  $("layout").classList.add("demo");
  $("demo-doc").hidden = false;
  var host = createDemoHost($("demo-content"));
  start(host);

  function open(file) {
    $("status").textContent = file.name + " wordt geopend…";
    host.openFile(file)
      .then(function (name) {
        $("demo-name").textContent = name;
        state.stack = [];
        state.expanded.clear();
        state.savedPlace = null;
        $("return").hidden = true;
        $("search").value = "";
        renderDetail();
        return rescan(false);
      })
      .catch(function (e) { $("status").textContent = e.message; });
  }

  $("demo-open").addEventListener("click", function () { $("demo-file").click(); });
  $("demo-file").addEventListener("change", function () {
    if (this.files[0]) open(this.files[0]);
    this.value = "";
  });
  var doc = $("demo-doc");
  doc.addEventListener("dragover", function (e) {
    e.preventDefault();
    doc.classList.add("dragging");
  });
  doc.addEventListener("dragleave", function () { doc.classList.remove("dragging"); });
  doc.addEventListener("drop", function (e) {
    e.preventDefault();
    doc.classList.remove("dragging");
    if (e.dataTransfer.files[0]) open(e.dataTransfer.files[0]);
  });
}

if (typeof Office !== "undefined") {
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) start(createWordHost());
    else startDemo();
  });
} else {
  startDemo();
}
