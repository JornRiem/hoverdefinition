const test = require("node:test");
const assert = require("node:assert/strict");
const D = require("../web/definitions.js");

const scheduleTable = [
  ["Term", "Meaning"],
  ["“Business Day”", "a day (other than a Saturday or Sunday) on which banks are open in Amsterdam."],
  ["Completion", "completion of the sale and purchase of the Shares in accordance with clause 6."],
  ["Completion Date", "the date falling five Business Days after the Condition is satisfied."],
  ["Shares", "all issued shares in the capital of the Company."]
];

test("leest een definitietabel met twee kolommen en slaat de kopregel over", () => {
  const defs = D.parseTable(scheduleTable, 0);
  assert.deepEqual(defs.map((d) => d.term), ["Business Day", "Completion", "Completion Date", "Shares"]);
  assert.deepEqual(defs[2].source, { kind: "table", table: 0, row: 3 });
});

test("negeert tabellen die geen definitietabel zijn", () => {
  const prices = [
    ["Item", "Price"],
    ["123", "EUR 10"],
    ["fee for services rendered under this agreement", "EUR 20"],
    ["other costs", "EUR 30"]
  ];
  assert.deepEqual(D.parseTable(prices, 0), []);
});

test("herkent Engelse en Nederlandse definitie-alinea's", () => {
  const cases = [
    ["1.1 “Agreement” means this share purchase agreement.", "Agreement", "means this share purchase agreement."],
    ["\"Encumbrance\": any mortgage, pledge or lien.", "Encumbrance", "any mortgage, pledge or lien."],
    ["(a) \"Losses\" has the meaning given in clause 9.", "Losses", "has the meaning given in clause 9."],
    ["“Koper” betekent Beta B.V.", "Koper", "betekent Beta B.V."],
    ["Onder “Aandelen” wordt verstaan: alle geplaatste aandelen.", "Aandelen", "alle geplaatste aandelen."]
  ];
  for (const [text, term, meaning] of cases) {
    const [def] = D.parseParagraph(text, 7);
    assert.equal(def.term, term, text);
    assert.equal(def.text, meaning, text);
  }
});

test("herkent inline definities tussen haakjes", () => {
  const text = "Alpha B.V. (the “Seller”) and Beta B.V. (hierna: “Koper”), each a “Party” and together the “Parties”.";
  assert.deepEqual(D.parseParagraph(text, 0).map((d) => d.term), ["Seller", "Koper", "Party", "Parties"]);
  const both = D.parseParagraph("Alpha and Beta (each a “Party” and together the “Parties”).", 0);
  assert.deepEqual(both.map((d) => d.term), ["Party", "Parties"]);
});

test("herkent inline definities zonder haakjes", () => {
  const en = D.parseParagraph("Alpha B.V. (the “Seller”) and Beta B.V., each a “Party” and together the “Parties”.", 0);
  assert.deepEqual(en.map((d) => d.term), ["Seller", "Party", "Parties"]);
  const nl = D.parseParagraph("Alpha B.V., hierna te noemen “Verkoper”, en Beta B.V.", 0);
  assert.deepEqual(nl.map((d) => d.term), ["Verkoper"]);
});

test("ziet gewone tekst tussen haakjes niet aan voor een definitie", () => {
  assert.deepEqual(D.parseParagraph("The Seller shall pay (subject to the cap in “clause 9”) all Losses.", 0), []);
});

function sampleIndex() {
  return D.buildIndex(
    D.extractDefinitions({
      tables: [scheduleTable],
      paragraphs: [{ text: "This agreement is between Alpha B.V. (the “Seller”) and Beta B.V.", index: 0 }]
    })
  );
}

test("vindt de term onder de cursor, langste term eerst", () => {
  const index = sampleIndex();
  const text = "On the Completion Date the Seller shall deliver the Shares.";
  assert.equal(D.findTermAt(text, text.indexOf("Date") + 1, index), "Completion Date");
  assert.equal(D.findTermAt(text, text.indexOf("Seller"), index), "Seller");
  assert.equal(D.findTermAt(text, text.indexOf("deliver"), index), null);
});

test("herkent meervouden, bezittelijke vormen en geen woorddelen", () => {
  const index = sampleIndex();
  assert.equal(D.findTermAt("within two Business Days", 15, index), "Business Day");
  assert.equal(D.findTermAt("the Seller’s obligations", 6, index), "Seller");
  assert.equal(D.findTermAt("the Sellers", 6, index), "Seller");
  assert.equal(D.findTermAt("Resellers", 4, index), null);
  assert.equal(D.findTermAt("Seller", 0, index), "Seller");
  const pair = "Seller/Shares";
  assert.equal(D.findTermAt(pair, pair.indexOf("Shares") + 1, index), "Shares");
  assert.equal(D.findTermAt("the seller", 6, index), null);
});

test("splitst een definitie in tekst en klikbare termen, zonder zichzelf", () => {
  const index = sampleIndex();
  const def = index.byTerm.get("Completion Date")[0];
  const parts = D.tokenize("the Completion Date falls five Business Days after Completion", index, "Completion Date");
  assert.deepEqual(parts.filter((p) => p.term).map((p) => p.term), ["Business Day", "Completion"]);
  assert.equal(parts.map((p) => p.text).join(""), "the Completion Date falls five Business Days after Completion");
  assert.ok(def.text.includes("Business Days"));
});

test("vouwt definities recursief uit en stopt bij kringverwijzingen", () => {
  const index = D.buildIndex(
    D.extractDefinitions({
      tables: [[
        ["“Completion Date”", "five Business Days after the Condition is satisfied."],
        ["“Business Day”", "a day on which banks are open."],
        ["“Condition”", "approval of the Transaction."],
        ["“Transaction”", "the transaction subject to the Condition."]
      ]]
    })
  );
  const def = index.byTerm.get("Completion Date")[0];
  const keys = D.expansionKeys(def, index, ["Completion Date"], "root");
  // Business Day, Condition, en daarbinnen Transaction; Condition binnen Transaction is een kring.
  assert.deepEqual(keys, ["root/1", "root/3", "root/3/1"]);
  assert.equal(D.expansionKeys(def, index, ["Completion Date"], "root", { depth: 1, total: 300 }).length, 2);
});

test("termen uit dezelfde inline alinea vouwen niet in elkaar uit", () => {
  const index = D.buildIndex(
    D.extractDefinitions({
      paragraphs: [{ text: "Alpha (the “Seller”) and Beta (the “Purchaser”), each a “Party”.", index: 0 }]
    })
  );
  const seller = index.byTerm.get("Seller")[0];
  assert.deepEqual(D.tokenize(seller.text, index, D.ownTerms(seller)).filter((p) => p.term), []);
  assert.deepEqual(D.expansionKeys(seller, index, ["Seller"], "root"), []);
});

test("gebruikt een geselecteerde term als de cursor niet op een term staat", () => {
  const index = sampleIndex();
  assert.equal(D.resolve("“Shares”", "", 0, index), "Shares");
  assert.equal(D.resolve("", "no terms here", 3, index), null);
});
