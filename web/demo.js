/*
 * Demo-modus: als de pagina buiten Word wordt geopend, toont de add-in een
 * voorbeeldcontract zodat je de werking in een gewone browser kunt bekijken.
 * Je kunt ook een eigen .docx openen; dat wordt alleen in de browser gelezen.
 */
(function () {
  "use strict";

  var BLOCKS = [
    { type: "h", text: "Share Purchase Agreement" },
    { type: "p", text: "This agreement is entered into between Alpha Holding B.V. (the “Seller”) and Beta Investments B.V. (the “Purchaser”), each a “Party” and together the “Parties”." },
    { type: "h", text: "2. Sale and purchase" },
    { type: "p", text: "2.1 On the terms of this Agreement, the Seller shall sell and the Purchaser shall purchase the Shares free from any Encumbrance with effect from Completion." },
    { type: "p", text: "2.2 The Purchase Price shall be paid by the Purchaser on the Completion Date to the Seller’s Account." },
    { type: "h", text: "6. Completion" },
    { type: "p", text: "6.1 Completion shall take place on the Completion Date at the offices of the Seller’s lawyers, or at such other place as the Parties may agree." },
    { type: "p", text: "6.2 If the Condition is not satisfied on or before the Long Stop Date, either Party may terminate this Agreement." },
    { type: "h", text: "Schedule 1 – Definitions" },
    {
      type: "table",
      rows: [
        ["Term", "Meaning"],
        ["“Account”", "the bank account of the Seller notified in writing to the Purchaser at least three Business Days before the Completion Date."],
        ["“Agreement”", "this share purchase agreement, including its Schedules."],
        ["“Business Day”", "a day (other than a Saturday or Sunday or a public holiday in the Netherlands) on which banks are open in Amsterdam for general business."],
        ["“Company”", "Gamma Technologies B.V., a private limited liability company with its seat in Utrecht."],
        ["“Completion”", "completion of the sale and purchase of the Shares in accordance with clause 6."],
        ["“Completion Date”", "the date falling five Business Days after the date on which the Condition is satisfied, or such other date as the Parties may agree in writing."],
        ["“Condition”", "the approval of the Transaction by the Authority for Consumers and Markets."],
        ["“Encumbrance”", "any pledge, right of usufruct, attachment or other security interest or third-party right of whatever nature."],
        ["“Long Stop Date”", "the date falling six months after the date of this Agreement, or if that day is not a Business Day, the next Business Day."],
        ["“Purchase Price”", "EUR 12,500,000, subject to adjustment in accordance with Schedule 3."],
        ["“Shares”", "all issued and outstanding shares in the capital of the Company."],
        ["“Transaction”", "the sale and purchase of the Shares contemplated by this Agreement."]
      ]
    }
  ];

  function createDemoHost(container) {
    var paragraphEls = [];
    var tableEls = [];
    var listeners = [];
    var lastPlace = null;

    function render(blocks) {
      container.replaceChildren();
      paragraphEls = [];
      tableEls = [];
      lastPlace = null;
      blocks.forEach(function (block) {
        if (block.type === "table") {
          var table = document.createElement("table");
          block.rows.forEach(function (row) {
            var tr = table.insertRow();
            row.forEach(function (cell) { tr.insertCell().textContent = cell; });
          });
          container.appendChild(table);
          tableEls.push(table);
        } else {
          var el = document.createElement(block.type === "h" ? "h2" : "p");
          el.textContent = block.text;
          container.appendChild(el);
          paragraphEls.push(el);
        }
      });
    }

    render(BLOCKS);

    container.addEventListener("mouseup", function (e) {
      var block = e.target.closest("p, h2, td");
      var sel = window.getSelection();
      if (!block || !sel.rangeCount) return;
      var range = sel.getRangeAt(0);
      if (!block.contains(range.startContainer)) return;
      var before = document.createRange();
      before.setStart(block, 0);
      before.setEnd(range.startContainer, range.startOffset);
      lastPlace = block;
      var info = { selected: sel.toString(), paragraph: block.textContent, offset: before.toString().length };
      listeners.forEach(function (cb) { cb(info); });
    });

    function flash(el) {
      el.scrollIntoView({ behavior: "smooth", block: "center" });
      el.classList.add("flash");
      setTimeout(function () { el.classList.remove("flash"); }, 1500);
    }

    return {
      scan: function () {
        return Promise.resolve({
          tables: tableEls.map(function (t) {
            return Array.from(t.rows).map(function (r) {
              return Array.from(r.cells).map(function (c) { return c.textContent; });
            });
          }),
          paragraphs: paragraphEls.map(function (el, i) { return { text: el.textContent, index: i }; })
        });
      },
      onSelection: function (cb) { listeners.push(cb); },
      goTo: function (def) {
        var s = def.source;
        var el = s.kind === "table" ? tableEls[s.table].rows[s.row].cells[0] : paragraphEls[s.paragraph];
        if (el) flash(el);
        return Promise.resolve();
      },
      returnToText: function (place) {
        if (place) flash(place);
        return Promise.resolve();
      },
      savePlace: function () { return Promise.resolve(lastPlace); },
      openFile: function (file) {
        return docxToBlocks(file).then(function (blocks) {
          render(blocks);
          container.scrollTop = 0;
          window.scrollTo(0, 0);
          return file.name;
        });
      }
    };
  }

  /* ---------- Eigen .docx inlezen ---------- */

  function loadMammoth() {
    if (window.mammoth) return Promise.resolve(window.mammoth);
    return new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = "vendor/mammoth.browser.min.js";
      s.onload = function () { resolve(window.mammoth); };
      s.onerror = function () { reject(new Error("De Word-lezer kon niet worden geladen. Probeer de pagina opnieuw te openen.")); };
      document.head.appendChild(s);
    });
  }

  function docxToBlocks(file) {
    if (!/\.docx$/i.test(file.name)) {
      return Promise.reject(new Error("Alleen .docx-bestanden worden ondersteund. Sla een .doc-bestand in Word eerst op als .docx."));
    }
    return Promise.all([loadMammoth(), file.arrayBuffer()])
      .then(function (r) { return r[0].convertToHtml({ arrayBuffer: r[1] }); })
      .then(function (result) {
        var blocks = htmlToBlocks(result.value);
        if (!blocks.length) throw new Error("Er is geen tekst gevonden in " + file.name + ".");
        return blocks;
      }, function (e) {
        if (e && /Word-lezer/.test(e.message)) throw e;
        throw new Error(file.name + " kon niet worden gelezen. Is het een geldig Word-document?");
      });
  }

  // Zet de HTML van mammoth om naar dezelfde blokken als het voorbeeldcontract.
  function htmlToBlocks(html) {
    var root = new DOMParser().parseFromString("<div>" + html + "</div>", "text/html").body.firstChild;
    var blocks = [];

    function text(el) {
      return el.textContent.replace(/\s+/g, " ").trim();
    }

    function cellText(cell) {
      var parts = Array.from(cell.querySelectorAll("p, li, h1, h2, h3, h4, h5, h6")).map(text).filter(Boolean);
      return parts.length ? parts.join("\n") : text(cell);
    }

    function walk(parent) {
      Array.from(parent.children).forEach(function (el) {
        var tag = el.tagName;
        if (tag === "TABLE") {
          blocks.push({
            type: "table",
            rows: Array.from(el.rows).map(function (row) { return Array.from(row.cells).map(cellText); })
          });
        } else if (/^H[1-6]$/.test(tag)) {
          if (text(el)) blocks.push({ type: "h", text: text(el) });
        } else if (tag === "P") {
          if (text(el)) blocks.push({ type: "p", text: text(el) });
        } else if (tag === "LI") {
          var own = Array.from(el.childNodes)
            .filter(function (n) { return !(n.nodeType === 1 && /^(UL|OL)$/.test(n.tagName)); })
            .map(function (n) { return n.textContent; })
            .join("").replace(/\s+/g, " ").trim();
          if (own) blocks.push({ type: "p", text: own });
          walk(el);
        } else {
          walk(el);
        }
      });
    }

    walk(root);
    return blocks;
  }

  window.createDemoHost = createDemoHost;
})();
