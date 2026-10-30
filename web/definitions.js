/*
 * Herkennen van gedefinieerde termen in contracttekst.
 *
 * Dit bestand bevat alleen pure logica (geen Word-API), zodat het zowel in de
 * add-in als in de tests (Node) gebruikt kan worden.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Definitions = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var Q_OPEN = "\"'“”‘’„«";
  var Q_CLOSE = "\"'“”‘’»";
  var QUOTED = new RegExp("[" + Q_OPEN + "]([^" + Q_OPEN + Q_CLOSE + "]{1,80})[" + Q_CLOSE + "]", "gu");

  var HEADER_WORDS = /^(defined\s+)?(term|terms|definition|definitions|word|expression|meaning|begrip|begrippen|term(en)?|definitie|definities|betekenis)$/i;

  // "Term" means ... / "Term": ... / "Term" betekent ...
  // Optionele nummering vooraan: 1.1, 1.1., (a), a), (iv)
  var LEADING_DEFINITION = new RegExp(
    "^\\s*(?:(?:\\d+(?:\\.\\d+)*\\.?|\\(?[a-z0-9]{1,4}[.)])\\s+)?" +
      "[" + Q_OPEN + "]([^" + Q_OPEN + Q_CLOSE + "]{1,80})[" + Q_CLOSE + "]" +
      "\\s*,?\\s*" +
      "((?:means|mean|shall mean|has the meaning|have the meaning|shall have the meaning|includes|include|shall include|refers to|is defined|" +
      "betekent|betekenen|heeft de betekenis|hebben de betekenis|houdt in|omvat|wordt verstaan)\\b|[:–—-])",
    "iu"
  );

  // onder "Koper" wordt verstaan: ...
  var DUTCH_ONDER = new RegExp(
    "^\\s*(?:(?:\\d+(?:\\.\\d+)*\\.?|\\(?[a-z0-9]{1,4}[.)])\\s+)?onder\\s+[" + Q_OPEN + "]([^" + Q_OPEN + Q_CLOSE + "]{1,80})[" + Q_CLOSE + "]\\s+(?:wordt|worden)\\s+verstaan\\s*:?\\s*",
    "iu"
  );

  // Woorden die in een inline definitie tussen haakjes mogen staan,
  // bijvoorbeeld (the "Seller") of (hierna: "Koper").
  var INLINE_FILLER = /^(?:the|a|an|each|together|collectively|jointly|individually|separately|hereinafter|hereafter|herein|referred|called|to|as|and|or|such|being|hierna|te|noemen|genoemd|de|het|een|elk|ieder|iedere|elke|gezamenlijk|gezamenlijke|tezamen|samen|respectievelijk|en|of|als)$/i;

  var INLINE_LEAD = new RegExp(
    "\\b(?:each|together|collectively|jointly|individually|hereinafter(?: referred to as)?|referred to as|hierna(?: te noemen| genoemd)?|gezamenlijk|tezamen)" +
      "[\\s:,]*(?:(?:a|an|the|de|het|een)\\s+)?[" + Q_OPEN + "]([^" + Q_OPEN + Q_CLOSE + "]{1,80})[" + Q_CLOSE + "]",
    "giu"
  );

  function normalizeSpace(s) {
    return String(s || "")
      .replace(/[\u0007\u000b]/g, " ")
      .replace(/\r\n?/g, "\n")
      .replace(/[ \t ]+/g, " ")
      .replace(/ *\n */g, "\n")
      .trim();
  }

  function cleanTerm(raw) {
    var s = normalizeSpace(raw).replace(/\n/g, " ");
    s = s.replace(new RegExp("^[\\s" + Q_OPEN + Q_CLOSE + "]+|[\\s" + Q_OPEN + Q_CLOSE + ":;,.]+$", "gu"), "");
    s = s.replace(new RegExp("[" + Q_OPEN + Q_CLOSE + "]", "gu"), "");
    return s.trim();
  }

  function looksLikeTerm(term) {
    if (!term || term.length > 80) return false;
    if (term.split(/\s+/).length > 8) return false;
    return /^[\p{Lu}\p{N}]/u.test(term);
  }

  function isHeaderRow(left, right) {
    return HEADER_WORDS.test(cleanTerm(left)) && (!right || HEADER_WORDS.test(cleanTerm(right)));
  }

  /*
   * Een tabel met (minstens) twee kolommen waarvan de linkerkolom termen bevat
   * en de rechterkolom de betekenis. Geeft [] terug als de tabel er niet op lijkt.
   */
  function parseTable(rows, tableIndex) {
    var candidates = [];
    var usable = 0;
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r] || [];
      if (row.length < 2) continue;
      var left = row[0];
      var right = row.slice(1).join("\n");
      if (!normalizeSpace(left) && !normalizeSpace(right)) continue;
      usable++;
      if (r === 0 && isHeaderRow(left, row[1])) {
        usable--;
        continue;
      }
      var term = cleanTerm(left);
      var text = normalizeSpace(right);
      if (looksLikeTerm(term) && text) {
        candidates.push({ term: term, text: text, source: { kind: "table", table: tableIndex, row: r } });
      }
    }
    if (candidates.length < 2 || candidates.length < usable * 0.6) return [];
    return candidates;
  }

  function parseParagraph(rawText, paragraphIndex) {
    var text = normalizeSpace(rawText);
    if (!text) return [];

    var m = DUTCH_ONDER.exec(text);
    if (m && looksLikeTerm(cleanTerm(m[1]))) {
      return [{ term: cleanTerm(m[1]), text: text.slice(m[0].length).trim(), source: { kind: "paragraph", paragraph: paragraphIndex } }];
    }

    m = LEADING_DEFINITION.exec(text);
    if (m && looksLikeTerm(cleanTerm(m[1]))) {
      var rest = text.slice(m.index + m[0].length - m[2].length);
      if (/^[:–—-]/.test(rest)) rest = rest.slice(1);
      rest = rest.trim();
      if (rest) {
        return [{ term: cleanTerm(m[1]), text: rest, source: { kind: "paragraph", paragraph: paragraphIndex } }];
      }
    }

    return parseInline(text, paragraphIndex);
  }

  // (the "Seller"), (each a "Party" and together the "Parties"), (hierna: "Koper")
  function parseInline(text, paragraphIndex) {
    var found = [];
    var paren = /\(([^()]{1,200})\)/g;
    var p;
    while ((p = paren.exec(text))) {
      var inner = p[1];
      var terms = [];
      var q;
      QUOTED.lastIndex = 0;
      while ((q = QUOTED.exec(inner))) terms.push(cleanTerm(q[1]));
      if (!terms.length) continue;
      var rest = inner.replace(QUOTED, " ").split(/[\s,:;]+/).filter(Boolean);
      if (!rest.every(function (w) { return INLINE_FILLER.test(w); })) continue;
      terms.forEach(function (term) {
        if (looksLikeTerm(term)) {
          found.push({ term: term, text: text, source: { kind: "inline", paragraph: paragraphIndex } });
        }
      });
    }
    // each a "Party" and together the "Parties", hierna te noemen "Koper"
    var m;
    INLINE_LEAD.lastIndex = 0;
    while ((m = INLINE_LEAD.exec(text))) {
      var term = cleanTerm(m[1]);
      var known = found.some(function (d) { return d.term === term; });
      if (!known && looksLikeTerm(term)) {
        found.push({ term: term, text: text, source: { kind: "inline", paragraph: paragraphIndex } });
      }
    }
    return found;
  }

  /*
   * Verzamelt alle definities uit een gescand document:
   *   { tables: [ [ [cel, cel], ... ], ... ], paragraphs: [ { text, index } ] }
   * Alinea's die in een tabel staan moeten al zijn weggefilterd.
   */
  function extractDefinitions(doc) {
    var defs = [];
    (doc.tables || []).forEach(function (rows, t) {
      defs = defs.concat(parseTable(rows, t));
    });
    (doc.paragraphs || []).forEach(function (p) {
      defs = defs.concat(parseParagraph(p.text, p.index));
    });
    return defs;
  }

  function escapeRegExp(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  // Meervouden en bezittelijke vormen, Engels en Nederlands.
  function variantsOf(term) {
    var v = [term, term + "s", term + "es", term + "'s", term + "’s", term + "en", term + "n"];
    if (/[^aeiou]y$/.test(term)) v.push(term.slice(0, -1) + "ies");
    if (/[^aeiou]ies$/.test(term)) v.push(term.slice(0, -3) + "y");
    if (/s$/.test(term) && term.length > 3) v.push(term.slice(0, -1));
    return v;
  }

  function buildIndex(defs) {
    var byTerm = new Map();
    defs.forEach(function (d) {
      if (!byTerm.has(d.term)) byTerm.set(d.term, []);
      var list = byTerm.get(d.term);
      if (!list.some(function (x) { return x.text === d.text; })) list.push(d);
    });
    // Tabeldefinities eerst, inline definities als laatste.
    var rank = { table: 0, paragraph: 1, inline: 2 };
    byTerm.forEach(function (list) {
      list.sort(function (a, b) { return rank[a.source.kind] - rank[b.source.kind]; });
    });

    // Varianten verwijzen naar de canonieke term; echte termen gaan altijd voor.
    var lookup = new Map();
    byTerm.forEach(function (_, term) { lookup.set(term, term); });
    byTerm.forEach(function (_, term) {
      variantsOf(term).forEach(function (v) {
        if (!lookup.has(v)) lookup.set(v, term);
      });
    });

    var forms = Array.from(lookup.keys()).sort(function (a, b) { return b.length - a.length; });
    var regex = forms.length
      ? new RegExp("(?<![\\p{L}\\p{N}])(?:" + forms.map(escapeRegExp).join("|") + ")(?![\\p{L}\\p{N}])", "gu")
      : null;

    return { byTerm: byTerm, lookup: lookup, regex: regex };
  }

  function matches(text, index) {
    var out = [];
    if (!index.regex || !text) return out;
    index.regex.lastIndex = 0;
    var m;
    while ((m = index.regex.exec(text))) {
      out.push({ start: m.index, end: m.index + m[0].length, form: m[0], term: index.lookup.get(m[0]) });
    }
    return out;
  }

  // De term waar de cursor (offset) in of direct tegenaan staat.
  function findTermAt(text, offset, index) {
    var all = matches(text, index);
    for (var i = 0; i < all.length; i++) {
      if (offset >= all[i].start && offset <= all[i].end) return all[i].term;
    }
    return null;
  }

  function resolve(selected, paragraph, offset, index) {
    var fromCursor = findTermAt(paragraph, offset, index);
    if (fromCursor) return fromCursor;
    var s = cleanTerm(selected);
    return (s && index.lookup.get(s)) || null;
  }

  // Splitst een definitie in stukken tekst en (klikbare) termen.
  function tokenize(text, index, exclude) {
    var parts = [];
    var pos = 0;
    matches(text, index).forEach(function (m) {
      if (m.term === exclude) return;
      if (m.start > pos) parts.push({ text: text.slice(pos, m.start) });
      parts.push({ text: m.form, term: m.term });
      pos = m.end;
    });
    if (pos < text.length) parts.push({ text: text.slice(pos) });
    return parts;
  }

  return {
    cleanTerm: cleanTerm,
    parseTable: parseTable,
    parseParagraph: parseParagraph,
    extractDefinitions: extractDefinitions,
    buildIndex: buildIndex,
    findTermAt: findTermAt,
    resolve: resolve,
    tokenize: tokenize
  };
});
