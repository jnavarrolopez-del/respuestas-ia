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

test("el mensaje a la IA lleva el contexto y las celdas", () => {
  const msg = L.buildUserMessage(
    [{ id: 0, sheet: "5_Estadistica", address: "G10", question: "Media aritmética (x̄)", kind: "auto", extra: "no" }],
    "### Hoja 'X'\nA1: hola"
  );
  assert.ok(msg.startsWith("CONTEXTO DEL LIBRO:\n### Hoja 'X'"));
  const cells = JSON.parse(msg.split("CELDAS:\n")[1]);
  assert.deepEqual(cells, [{ id: 0, hoja: "5_Estadistica", celda: "G10", texto_pegado: "Media aritmética (x̄)", expected: "auto" }]);
  assert.ok(!L.buildUserMessage([{ id: 1, question: "¿2+2?", kind: "number" }]).includes("CONTEXTO"));
});

test("columnas y direcciones de celda", () => {
  assert.equal(L.columnLetter(0), "A");
  assert.equal(L.columnLetter(25), "Z");
  assert.equal(L.columnLetter(26), "AA");
  assert.equal(L.columnLetter(701), "ZZ");
  assert.equal(L.columnLetter(702), "AAA");
  assert.equal(L.cellAddress(9, 6), "G10");
});

test("collectTexts acepta etiquetas y encabezados, no fórmulas ni números", () => {
  assert.deepEqual(L.collectTexts([["Media aritmética (x̄)", 4.91, "=A1", "", "x"], ["Índice en serie"]]), [
    "Media aritmética (x̄)",
    "Índice en serie",
  ]);
});

test("una fórmula de la IA tiene prioridad y se valida", () => {
  const ai = (formula, answer = "") => ({ status: "ok", formula, answer, option_label: "", option_text: "" });
  assert.deepEqual(L.coerceAnswer(ai("=AVERAGE(B10:B19)"), { kind: "number" }), { ok: true, formula: "=AVERAGE(B10:B19)" });
  assert.deepEqual(L.coerceAnswer(ai("  ='2_Tasas'!D11 "), { kind: "auto" }), { ok: true, formula: "='2_Tasas'!D11" });
  // Sin "=" o con saltos de línea no es una fórmula válida: se usa answer.
  assert.deepEqual(L.coerceAnswer(ai("AVERAGE(B1:B2)", "4"), { kind: "auto" }), { ok: true, value: 4 });
  assert.deepEqual(L.coerceAnswer(ai("=A1\n+A2", ""), { kind: "auto" }), { ok: false, reason: "unknown" });
  assert.deepEqual(L.coerceAnswer({ status: "unknown", formula: "=A1", answer: "" }, { kind: "auto" }), { ok: false, reason: "unknown" });
});

test("buildWorkbookContext: hoja activa primero, celdas a rellenar marcadas, recorte", () => {
  const sheets = [
    { name: "Datos", rowIndex: 0, columnIndex: 0, formulas: [["Año", "Café"], [2014, 8.2]] },
    {
      name: "5_Estadistica",
      rowIndex: 8,
      columnIndex: 5,
      formulas: [["Medida", "Valor"], ["Media", ""], ["Mediana", ""]],
      marks: [[false, false], [false, true], [false, false]],
    },
  ];
  const ctx = L.buildWorkbookContext(sheets, "5_Estadistica");
  assert.ok(ctx.indexOf("### Hoja '5_Estadistica' (activa)") < ctx.indexOf("### Hoja 'Datos'"));
  assert.ok(ctx.includes("F9: Medida"));
  assert.ok(ctx.includes("G10:  [RELLENAR]"));
  assert.ok(!ctx.includes("G11"));
  assert.ok(ctx.includes("B2: 8.2"));
  const withValues = L.buildWorkbookContext(
    [{ name: "T", rowIndex: 9, columnIndex: 0, formulas: [["='2_Tasas'!A11", "=1/3"]], values: [[2015, 0.3333333333333333]] }],
    "T"
  );
  assert.ok(withValues.includes("A10: ='2_Tasas'!A11 → 2015"));
  assert.ok(withValues.includes("B10: =1/3 → 0.3333333333"));
  const small = L.buildWorkbookContext(sheets, "Datos", { maxChars: 260 });
  assert.ok(small.length <= 300);
});

test("pickSolveTargets: celdas marcadas primero, sin celdas tapadas por combinaciones", () => {
  const f = [
    ["Año", "Precio", "", ""],
    [2014, 8.2, "", ""],
    [2015, 8.45, "", "x"],
  ];
  const marked = (r, c) => r > 0 && c === 2;
  assert.deepEqual(L.pickSolveTargets(f, marked, () => false), [{ r: 1, c: 2 }, { r: 2, c: 2 }]);
  // Sin marcas: todas las vacías, salvo las tapadas.
  assert.deepEqual(L.pickSolveTargets(f, () => false, (r, c) => c === 3), [{ r: 0, c: 2 }, { r: 1, c: 2 }, { r: 2, c: 2 }]);
  // Sin marcas y demasiadas vacías: nada (no rellenar media hoja por error).
  assert.deepEqual(L.pickSolveTargets(f, () => false, () => false, 2), []);
});

test("solo los errores de fórmula mal escrita invalidan la respuesta", () => {
  for (const e of ["#NAME?", "#¿NOMBRE?", "#REF!", "#NULL!", "#¡NULO!"]) assert.ok(L.isBrokenFormulaError(e), e);
  for (const e of ["#DIV/0!", "#N/A", "#VALUE!", "#¡VALOR!", "#¡DIV/0!", 12]) assert.ok(!L.isBrokenFormulaError(e), String(e));
});

test("el aviso a la IA explica el modo resolver (texto vacío)", () => {
  assert.ok(L.SYSTEM_PROMPT.includes("texto_pegado está vacío"));
});

test("el esquema de respuesta exige la fórmula", () => {
  assert.ok(L.ANSWER_SCHEMA.properties.answers.items.required.includes("formula"));
});
