// Pruebas de la lógica pura. Ejecutar con:  node --test test/
const test = require("node:test");
const assert = require("node:assert/strict");
const L = require("../src/logic.js");

test("detecta preguntas y descarta encabezados", () => {
  assert.ok(L.looksLikeQuestion("¿Cuál es la capital de Francia?"));
  assert.ok(L.looksLikeQuestion("Calcula el 15% de 200"));
  assert.ok(L.looksLikeQuestion("What is 2 + 2"));
  assert.ok(L.looksLikeQuestion("Capital de Italia: a) Roma b) Milán c) Turín"));
  assert.ok(!L.looksLikeQuestion("Nombre"));
  assert.ok(!L.looksLikeQuestion("Respuesta"));
  assert.ok(!L.looksLikeQuestion("Total ventas"));
  assert.ok(!L.looksLikeQuestion("=SUMA(A1:A3)"));
  assert.ok(!L.looksLikeQuestion(42));
  assert.ok(!L.looksLikeQuestion(""));
});

test("collectQuestions mantiene el orden y quita duplicados", () => {
  const v = [["Pregunta", "Respuesta"], ["¿2+2?", ""], ["¿Capital de España?", ""], ["¿2+2?", ""]];
  assert.deepEqual(L.collectQuestions(v), ["¿2+2?", "¿Capital de España?"]);
});

test("tipo esperado según formato, lista y encabezado", () => {
  assert.equal(L.inferExpectedType({ numberFormat: "0.00" }).kind, "number");
  assert.equal(L.inferExpectedType({ numberFormat: "#,##0 €" }).kind, "number");
  assert.equal(L.inferExpectedType({ numberFormat: "0%" }).kind, "percent");
  assert.equal(L.inferExpectedType({ numberFormat: "dd/mm/yyyy" }).kind, "date");
  assert.equal(L.inferExpectedType({ numberFormat: "@" }).kind, "text");
  assert.equal(L.inferExpectedType({ numberFormat: "General" }).kind, "auto");
  assert.equal(L.inferExpectedType({ numberFormat: "General", header: "Cantidad" }).kind, "number");
  assert.equal(L.inferExpectedType({ numberFormat: "General", header: "Año" }).kind, "number");
  const opt = L.inferExpectedType({ numberFormat: "General", listOptions: ["A", "B", "C"] });
  assert.equal(opt.kind, "option");
  assert.deepEqual(opt.options, ["A", "B", "C"]);
});

test("parseNumber entiende formatos españoles e ingleses", () => {
  assert.equal(L.parseNumber("1.234,5"), 1234.5);
  assert.equal(L.parseNumber("1,234.5"), 1234.5);
  assert.equal(L.parseNumber("3,5"), 3.5);
  assert.equal(L.parseNumber("42"), 42);
  assert.equal(L.parseNumber("-7"), -7);
  assert.equal(L.parseNumber("25%"), 0.25);
  assert.equal(L.parseNumber("1.000.000"), 1000000);
  assert.equal(L.parseNumber("unos 40"), null);
  assert.equal(L.parseNumber("París"), null);
});

test("coerceAnswer respeta el tipo de la celda", () => {
  const ok = (answer, extra = {}) => ({ status: "ok", answer, option_label: "", option_text: "", ...extra });
  assert.deepEqual(L.coerceAnswer(ok("30"), { kind: "number" }), { ok: true, value: 30 });
  assert.deepEqual(L.coerceAnswer(ok("treinta"), { kind: "number" }), { ok: false, reason: "mismatch" });
  assert.deepEqual(L.coerceAnswer(ok("25%"), { kind: "percent" }), { ok: true, value: 0.25 });
  assert.deepEqual(L.coerceAnswer(ok("25"), { kind: "percent" }), { ok: true, value: 0.25 });
  assert.deepEqual(L.coerceAnswer(ok("1492-10-12"), { kind: "date" }), { ok: true, value: "1492-10-12" });
  assert.deepEqual(L.coerceAnswer(ok("París"), { kind: "auto" }), { ok: true, value: "París" });
  assert.deepEqual(L.coerceAnswer(ok("4"), { kind: "auto" }), { ok: true, value: 4 });
  assert.deepEqual(L.coerceAnswer(ok("Roma", { option_label: "a", option_text: "Roma" }), { kind: "auto" }), { ok: true, value: "a" });
  assert.deepEqual(L.coerceAnswer(ok("=1+1"), { kind: "text" }), { ok: true, value: "'=1+1" });
  assert.deepEqual(L.coerceAnswer({ status: "unknown", answer: "" }, { kind: "auto" }), { ok: false, reason: "unknown" });
  assert.deepEqual(L.coerceAnswer({ status: "not_a_question", answer: "" }, { kind: "auto" }), { ok: false, reason: "not_a_question" });
  assert.deepEqual(L.coerceAnswer(null, { kind: "auto" }), { ok: false, reason: "unknown" });
});

test("coerceAnswer elige la opción del desplegable", () => {
  const ai = { status: "ok", answer: "b) Roma", option_label: "b", option_text: "Roma" };
  assert.deepEqual(L.coerceAnswer(ai, { kind: "option", options: ["Milán", "Roma", "Turín"] }), { ok: true, value: "Roma" });
  assert.deepEqual(L.coerceAnswer(ai, { kind: "option", options: ["a", "b", "c"] }), { ok: true, value: "b" });
  assert.deepEqual(L.coerceAnswer(ai, { kind: "option", options: ["A) Milán", "B) Roma"] }), { ok: true, value: "B) Roma" });
  const yes = { status: "ok", answer: "Verdadero", option_label: "", option_text: "" };
  assert.deepEqual(L.coerceAnswer(yes, { kind: "option", options: ["Verdadero", "Falso"] }), { ok: true, value: "Verdadero" });
  const none = { status: "ok", answer: "Nápoles", option_label: "", option_text: "" };
  assert.deepEqual(L.coerceAnswer(none, { kind: "option", options: ["Milán", "Roma"] }), { ok: false, reason: "mismatch" });
});

test("findPastedQuestions: solo celdas antes vacías y en orden", () => {
  const source = new Set(["¿Q1?", "¿Q2?", "¿Q3?"]);
  const after = [["¿Q1?"], ["¿Q2?"], ["¿Q3?"]];
  const before = [[""], ["texto propio"], [null]];
  const res = L.findPastedQuestions(after, (r, c) => ({ known: true, value: before[r][c] }), source, () => false);
  assert.deepEqual(res, [
    { r: 0, c: 0, question: "¿Q1?" },
    { r: 2, c: 0, question: "¿Q3?" },
  ]);
});

test("findPastedQuestions: ignora contenido desconocido, ajeno o escrito a mano", () => {
  const source = new Set(["¿Q1?"]);
  assert.deepEqual(L.findPastedQuestions([["¿Q1?"]], () => ({ known: false }), source, () => false), []);
  assert.deepEqual(L.findPastedQuestions([["¿Otra?"]], () => ({ known: true, value: "" }), source, () => false), []);
  assert.deepEqual(L.findPastedQuestions([["¿Q1?"]], () => ({ known: true, value: "" }), new Set(), () => false), []);
});

test("findPastedQuestions: reintento sobre una celda que falló", () => {
  const source = new Set(["¿Q1?"]);
  const res = L.findPastedQuestions([["¿Q1?"]], () => ({ known: true, value: "¿Q1?" }), source, () => true);
  assert.deepEqual(res, [{ r: 0, c: 0, question: "¿Q1?" }]);
});

test("el mensaje a la IA solo lleva la pregunta y el tipo", () => {
  const msg = JSON.parse(L.buildUserMessage([{ id: 0, question: "¿2+2?", kind: "number", extra: "secreto" }]));
  assert.deepEqual(msg, [{ id: 0, question: "¿2+2?", expected: "number" }]);
});
