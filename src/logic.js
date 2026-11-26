/*
 * Lógica pura (sin Office ni red) del complemento Respuestas IA.
 * Se usa en el navegador (window.RespuestasLogic) y en las pruebas con Node.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.RespuestasLogic = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const MAX_QUESTION_LENGTH = 4000;

  function normalize(s) {
    return String(s == null ? "" : s)
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function isEmpty(v) {
    return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  }

  const QUESTION_START = new RegExp(
    "^(?:\\d+[.)]\\s*)?(?:" +
      [
        // español
        "que", "cual", "cuales", "cuanto", "cuanta", "cuantos", "cuantas", "como", "donde", "adonde",
        "cuando", "quien", "quienes", "por que", "para que", "calcula", "calcule", "indica", "indique",
        "di", "nombra", "menciona", "define", "defina", "explica", "elige", "elija", "selecciona",
        "seleccione", "señala", "senala", "escribe", "escriba", "traduce", "traduzca", "convierte", "resuelve",
        "resuelva", "halla", "determina", "es verdad", "verdadero o falso", "v o f",
        // inglés
        "what", "which", "who", "whom", "whose", "when", "where", "why", "how", "is", "are", "does", "do",
        "did", "can", "calculate", "define", "choose", "select", "translate", "solve", "find",
        "true or false",
      ].join("|") +
      ")\\b",
    "i"
  );

  const OPTION_MARKERS = /(?:^|\s|\()(?:[a-eA-E]|[1-5])[).]\s*\S/g;

  /** ¿El texto de una celda parece una pregunta? Heurística local y conservadora. */
  function looksLikeQuestion(value) {
    if (typeof value !== "string") return false;
    const t = value.trim();
    if (t.length < 3 || t.length > MAX_QUESTION_LENGTH) return false;
    if (t.startsWith("=")) return false;
    if (/[?¿]/.test(t)) return true;
    // Sin signo de interrogación exigimos al menos tres palabras para no confundir encabezados.
    if (QUESTION_START.test(normalize(t)) && t.split(/\s+/).length >= 3) return true;
    const markers = t.match(OPTION_MARKERS);
    return !!markers && markers.length >= 2;
  }

  /** Textos de pregunta distintos que hay en una matriz de valores (en orden). */
  function collectQuestions(values) {
    const out = [];
    const seen = new Set();
    for (const row of values || []) {
      for (const v of row || []) {
        if (looksLikeQuestion(v) && !seen.has(v)) {
          seen.add(v);
          out.push(v);
        }
      }
    }
    return out;
  }

  function isDateFormat(fmt) {
    // Quita literales entre comillas y secuencias escapadas antes de buscar d/m/y.
    const f = String(fmt).replace(/"[^"]*"/g, "").replace(/\\./g, "").replace(/\[[^\]]*\]/g, "");
    if (/general/i.test(f)) return false;
    return /[dy]/i.test(f) || /(^|[^0#?])m{1,4}([^0#?]|$)/i.test(f) && !/[0#]/.test(f);
  }

  function isNumberFormat(fmt) {
    const f = String(fmt).replace(/"[^"]*"/g, "");
    if (f === "@") return false;
    return /[0#?]/.test(f) || /%/.test(f);
  }

  const NUMERIC_HEADER = /\b(n[ºo°]|num(?:ero)?|cantidad|total|suma|importe|precio|coste|costo|edad|ano|anos|year|years|valor|resultado num|puntos|nota|count|amount|price|cost|age|number|qty|cuant[oa]s?)\b/;

  /**
   * Tipo de respuesta esperado en la celda de destino, decidido en local.
   * snapshot: { numberFormat, listOptions: string[]|null, header: string|null }
   * Devuelve { kind: "option"|"number"|"date"|"percent"|"text"|"auto", options? }
   */
  function inferExpectedType(snapshot) {
    const s = snapshot || {};
    if (Array.isArray(s.listOptions) && s.listOptions.length > 0) {
      return { kind: "option", options: s.listOptions.slice() };
    }
    const fmt = s.numberFormat == null ? "General" : String(s.numberFormat);
    if (fmt === "@") return { kind: "text" };
    if (/%/.test(fmt.replace(/"[^"]*"/g, ""))) return { kind: "percent" };
    if (isDateFormat(fmt)) return { kind: "date" };
    if (isNumberFormat(fmt)) return { kind: "number" };
    if (s.header && NUMERIC_HEADER.test(normalize(s.header))) return { kind: "number" };
    return { kind: "auto" };
  }

  /** Instrucción de formato que se envía a la IA (no contiene datos de la hoja). */
  function typeInstruction(kind) {
    switch (kind) {
      case "number":
        return "number";
      case "percent":
        return "percent";
      case "date":
        return "date";
      case "option":
        return "option";
      case "text":
        return "short_text";
      default:
        return "auto";
    }
  }

  /** Convierte "1.234,5", "1,234.5", "3,5", "12 %", "-7" en número. Devuelve null si no es un número limpio. */
  function parseNumber(raw) {
    if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
    if (typeof raw !== "string") return null;
    let t = raw.trim().replace(/[\s  ]/g, "");
    let percent = false;
    if (t.endsWith("%")) {
      percent = true;
      t = t.slice(0, -1);
    }
    t = t.replace(/^[€$£]|[€$£]$/g, "");
    if (!/^[+-]?[\d.,]+(?:e[+-]?\d+)?$/i.test(t) || !/\d/.test(t)) return null;
    const lastDot = t.lastIndexOf(".");
    const lastComma = t.lastIndexOf(",");
    if (lastDot !== -1 && lastComma !== -1) {
      if (lastComma > lastDot) t = t.replace(/\./g, "").replace(",", ".");
      else t = t.replace(/,/g, "");
    } else if (lastComma !== -1) {
      const parts = t.split(",");
      // "1,234,567" = miles; "3,5" o "1,234" (ambiguo) = decimal al estilo español
      if (parts.length > 2) t = parts.join("");
      else t = t.replace(",", ".");
    } else if (lastDot !== -1) {
      const parts = t.split(".");
      if (parts.length > 2) t = parts.join("");
    }
    const n = Number(t);
    if (!Number.isFinite(n)) return null;
    return percent ? n / 100 : n;
  }

  function parseDate(raw) {
    if (typeof raw !== "string") return null;
    const t = raw.trim();
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    m = t.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    return null;
  }

  /** Elige la opción de una lista desplegable que corresponde a la respuesta de la IA. */
  function matchOption(ai, options) {
    const opts = (options || []).map((o) => ({ raw: o, n: normalize(o) })).filter((o) => o.n);
    if (!opts.length) return null;
    const candidates = [ai.answer, ai.option_text, ai.option_label]
      .filter((c) => typeof c === "string" && c.trim())
      .map(normalize);
    for (const c of candidates) {
      const hit = opts.find((o) => o.n === c);
      if (hit) return hit.raw;
    }
    const stripLabel = (s) => s.replace(/^\(?[a-z0-9]{1,2}[).:-]\s*/, "").trim();
    for (const c of candidates) {
      const sc = stripLabel(c);
      const hit = opts.find((o) => stripLabel(o.n) === sc && sc);
      if (hit) return hit.raw;
    }
    const label = normalize(ai.option_label || "").replace(/[^a-z0-9]/g, "");
    if (label) {
      const hit = opts.find((o) => {
        const m = o.n.match(/^\(?([a-z0-9]{1,2})[).:-]/);
        return m && m[1] === label;
      });
      if (hit) return hit.raw;
    }
    // Coincidencia parcial solo si es única.
    for (const c of candidates) {
      const sc = stripLabel(c);
      if (sc.length < 2) continue;
      const hits = opts.filter((o) => o.n.includes(sc) || sc.includes(stripLabel(o.n)));
      if (hits.length === 1) return hits[0].raw;
    }
    return null;
  }

  /**
   * ¿Un error de Excel indica una fórmula mal escrita (y no solo datos que aún faltan)?
   * #¿NOMBRE?/#NAME?, #REF!, #NULL!: fórmula incorrecta. #DIV/0!, #N/A, #VALUE!: puede faltar rellenar otras celdas.
   */
  function isBrokenFormulaError(value) {
    const v = normalize(value).toUpperCase();
    return /#\W*(NAME|NOMBRE|REF|NULL|NULO)\b/.test(v);
  }

  /**
   * Celdas a resolver dentro de una selección.
   * formulas: 2D del área; marked(r,c): celda con relleno claro; covered(r,c): parte no principal de una celda combinada.
   * Si hay celdas vacías marcadas, solo esas; si no, las vacías (como mucho maxUnmarked).
   */
  function pickSolveTargets(formulas, marked, covered, maxUnmarked) {
    const empties = [];
    const marks = [];
    (formulas || []).forEach((row, r) =>
      (row || []).forEach((v, c) => {
        if (!isEmpty(v) || (covered && covered(r, c))) return;
        empties.push({ r, c });
        if (marked && marked(r, c)) marks.push({ r, c });
      })
    );
    if (marks.length) return marks;
    return empties.length <= (maxUnmarked == null ? 60 : maxUnmarked) ? empties : [];
  }

  /** Acepta solo una fórmula de una línea que empiece por "=". Devuelve "" si no es válida. */
  function cleanFormula(f) {
    if (typeof f !== "string") return "";
    const t = f.trim();
    if (!t.startsWith("=") || t.length < 2 || t.length > 2000 || /[\r\n]/.test(t)) return "";
    return t;
  }

  /** Protege texto para que Excel no lo interprete como fórmula. */
  function safeText(s) {
    const t = String(s).trim();
    return /^[=+\-@]/.test(t) ? "'" + t : t;
  }

  /**
   * Convierte la respuesta de la IA al valor final de la celda según el tipo esperado.
   * ai: { status, answer, option_label, option_text }
   * Devuelve { ok: true, value } o { ok: false, reason: "not_a_question"|"unknown"|"mismatch" }.
   */
  function coerceAnswer(ai, expected) {
    if (!ai || typeof ai !== "object") return { ok: false, reason: "unknown" };
    if (ai.status === "not_a_question") return { ok: false, reason: "not_a_question" };
    if (ai.status !== "ok") return { ok: false, reason: "unknown" };
    const formula = cleanFormula(ai.formula);
    if (formula) return { ok: true, formula };
    const answer = typeof ai.answer === "string" ? ai.answer.trim() : String(ai.answer ?? "").trim();
    const kind = (expected && expected.kind) || "auto";

    if (kind === "option") {
      const hit = matchOption({ ...ai, answer }, expected.options);
      return hit == null ? { ok: false, reason: "mismatch" } : { ok: true, value: hit };
    }
    if (!answer) return { ok: false, reason: "unknown" };

    if (kind === "number") {
      const n = parseNumber(answer);
      return n == null ? { ok: false, reason: "mismatch" } : { ok: true, value: n };
    }
    if (kind === "percent") {
      let n = parseNumber(answer);
      if (n == null) return { ok: false, reason: "mismatch" };
      // "25" sin % en una celda de porcentaje significa 25 %.
      if (!/%/.test(answer) && Math.abs(n) > 1) n = n / 100;
      return { ok: true, value: n };
    }
    if (kind === "date") {
      const d = parseDate(answer);
      return d == null ? { ok: false, reason: "mismatch" } : { ok: true, value: d };
    }
    if (kind === "auto") {
      const n = parseNumber(answer);
      if (n != null && !/%/.test(answer)) return { ok: true, value: n };
      const label = typeof ai.option_label === "string" ? ai.option_label.trim() : "";
      if (label) return { ok: true, value: safeText(label) };
    }
    // Texto breve: una sola línea.
    const text = answer.split(/\r?\n/)[0].trim();
    if (!text) return { ok: false, reason: "unknown" };
    return { ok: true, value: safeText(text) };
  }

  /**
   * Decide qué celdas de un rango recién cambiado son "pegados de una pregunta copiada".
   * afterValues: matriz de valores tras el pegado.
   * beforeLookup(r, c): { known: boolean, value } — contenido previo de la celda.
   * sourceQuestions: Set de textos de la última selección con preguntas.
   * retryLookup(r, c, value): true si esa celda falló antes con esa misma pregunta.
   * Devuelve [{ r, c, question }] en orden de filas y columnas.
   */
  function findPastedQuestions(afterValues, beforeLookup, sourceQuestions, retryLookup) {
    const out = [];
    if (!sourceQuestions || sourceQuestions.size === 0) return out;
    (afterValues || []).forEach((row, r) => {
      (row || []).forEach((v, c) => {
        if (typeof v !== "string" || !sourceQuestions.has(v)) return;
        const before = beforeLookup(r, c);
        if (!before || !before.known) return;
        const emptyBefore = isEmpty(before.value);
        const retry = !emptyBefore && retryLookup && retryLookup(r, c, v);
        if (emptyBefore || retry) out.push({ r, c, question: v });
      });
    });
    return out;
  }

  /** Esquema JSON de la respuesta de la IA (salida estructurada). */
  const ANSWER_SCHEMA = {
    type: "object",
    properties: {
      answers: {
        type: "array",
        items: {
          type: "object",
          properties: {
            id: { type: "integer" },
            status: { type: "string", enum: ["ok", "not_a_question", "unknown"] },
            answer: { type: "string" },
            option_label: { type: "string" },
            option_text: { type: "string" },
            formula: { type: "string" },
          },
          required: ["id", "status", "formula", "answer", "option_label", "option_text"],
          additionalProperties: false,
        },
      },
    },
    required: ["answers"],
    additionalProperties: false,
  };

  const SYSTEM_PROMPT = [
    "Rellenas celdas de un libro de Excel. La persona copia un texto (una pregunta, una etiqueta como «Media aritmética» o el encabezado de una columna) y lo pega en una celda vacía: tú decides qué va en esa celda.",
    "Recibes el CONTEXTO del libro (cada hoja con sus celdas no vacías: dirección, valor o fórmula, y las celdas marcadas [RELLENAR] que el libro señala para completar) y una lista de CELDAS {id, hoja, celda, texto_pegado, expected}. Devuelve un elemento por cada id.",
    "",
    "Cuándo usar formula:",
    "- Si el valor de la celda se puede calcular con datos del libro, devuelve formula y deja answer vacío. Es lo preferido: el libro debe recalcularse si cambian los datos.",
    "- Si texto_pegado está vacío, la persona ha seleccionado la celda y ha pedido «resolver»: deduce qué va ahí por su posición, por las instrucciones de la hoja y por el enunciado. Si en esa celda no debe ir nada, status = not_a_question.",
    "- Cuando rellenes varias celdas de una misma tabla, usa fórmulas coherentes entre filas (la misma fórmula desplazada) y referencias entre ellas cuando el método lo pida (por ejemplo, la varianza a partir de la suma de la columna auxiliar).",
    "- Usa la posición de la celda para saber qué calcular: su fila (año, etiqueta de la columna A o de la columna de la izquierda) y su columna (encabezado). Si texto_pegado es un encabezado y se pegó en varias filas, cada fila lleva su propia fórmula.",
    "- Sigue las instrucciones escritas en la propia hoja (por ejemplo «En C10 escribe =B10/B$10*100») y respeta su método: si piden calcular algo «a mano», hazlo con las celdas auxiliares de la hoja, no con una función que lo resuelva de golpe.",
    "- Referencia celdas (también de otras hojas, como '2_Tasas'!D11 o Datos!B5); no escribas números que ya están en el libro.",
    "- Sintaxis: fórmula de Excel en inglés, separador de argumentos coma, punto decimal, empezando por =. Ejemplos: =AVERAGE(B10:B19), =QUARTILE(B10:B19,1), =VAR.P(B10:B19), =STDEV.P(B10:B19), =IF(G12>G10+G21,\"Sí\",\"No\"), =INDEX(A10:A19,MATCH(G12,B10:B19,0)). Pon entre comillas simples los nombres de hoja que empiezan por número o tienen espacios.",
    "- Para preguntas de interpretación que piden texto con cifras del libro, puedes devolver una fórmula que construya el texto (=\"…\"&TEXT(B5,\"0.00\")&\"…\") o un texto breve en answer.",
    "",
    "Cuándo usar answer (formula vacía):",
    "- Preguntas de conocimiento general que no dependen del libro.",
    "- answer contiene solo el valor final: sin explicaciones, sin unidades salvo que la pregunta las pida, sin comillas, sin punto final, una sola línea.",
    "- expected = number: answer es solo el número, con punto decimal y sin separador de miles (ej. 1234.5).",
    "- expected = percent: answer es el número seguido de % (ej. 25%).",
    "- expected = date: answer en formato AAAA-MM-DD.",
    "- expected = short_text: unas pocas palabras como máximo.",
    "- expected = option o la pregunta es de tipo test: answer es el texto de la opción correcta tal como aparece; option_label es su letra o número si las opciones lo tienen (ej. b), si no cadena vacía; option_text es el texto de la opción sin la letra.",
    "- expected = auto: si la respuesta natural es un número, da solo el número; si es tipo test, rellena option_label y option_text; si no, texto breve.",
    "",
    "Estado:",
    "- Si el texto pegado no pide nada (por ejemplo, la persona solo está copiando un nombre o un dato a otra celda), status = not_a_question y no rellenes nada.",
    "- Si no sabes qué poner con seguridad o falta información, status = unknown. No inventes.",
    "- En cualquier otro caso status = ok. Deja vacíos (cadena vacía) los campos que no apliquen.",
    "- Responde en el idioma de la pregunta.",
  ].join("\n");

  /** Mensaje para la IA: contexto del libro (si lo hay) y las celdas a rellenar. */
  function buildUserMessage(items, context) {
    const cells = items.map((it) => ({
      id: it.id,
      hoja: it.sheet || "",
      celda: it.address || "",
      texto_pegado: it.question,
      expected: typeInstruction(it.kind),
    }));
    const parts = [];
    if (context) parts.push("CONTEXTO DEL LIBRO:\n" + context);
    parts.push("CELDAS:\n" + JSON.stringify(cells));
    return parts.join("\n\n");
  }

  /** Letra de columna a partir del índice (0 -> A, 26 -> AA). */
  function columnLetter(index) {
    let n = index + 1;
    let s = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function cellAddress(row, col) {
    return columnLetter(col) + (row + 1);
  }

  /** Textos que se pueden copiar como "pregunta": cualquier texto escrito (no fórmulas). */
  function collectTexts(values) {
    const out = [];
    const seen = new Set();
    for (const row of values || []) {
      for (const v of row || []) {
        if (typeof v !== "string") continue;
        const t = v.trim();
        if (t.length < 2 || t.length > MAX_QUESTION_LENGTH || t.startsWith("=") || seen.has(v)) continue;
        seen.add(v);
        out.push(v);
      }
    }
    return out;
  }

  /**
   * Resume el libro como texto para la IA.
   * sheets: [{ name, rowIndex, columnIndex, formulas: 2D, values?: 2D, marks?: 2D booleans (celda a rellenar) }]
   * La hoja activa va primero y completa; el resto se recorta si el total supera maxChars.
   */
  function buildWorkbookContext(sheets, activeName, opts) {
    const o = Object.assign({ maxRows: 120, maxCols: 26, maxCell: 300, maxChars: 80000 }, opts || {});
    const ordered = [...sheets].sort((a, b) => (a.name === activeName ? -1 : b.name === activeName ? 1 : 0));
    const blocks = [];
    let total = 0;
    for (const sh of ordered) {
      const lines = [`### Hoja '${sh.name}'${sh.name === activeName ? " (activa)" : ""}`];
      const rows = Math.min((sh.formulas || []).length, o.maxRows);
      let truncated = (sh.formulas || []).length > o.maxRows;
      for (let r = 0; r < rows; r++) {
        const row = sh.formulas[r] || [];
        if (row.length > o.maxCols) truncated = true;
        for (let c = 0; c < Math.min(row.length, o.maxCols); c++) {
          const v = row[c];
          const mark = sh.marks && sh.marks[r] && sh.marks[r][c];
          const empty = v === "" || v === null || v === undefined;
          if (empty && !mark) continue;
          let s = empty ? "" : String(v).replace(/\s+/g, " ");
          // Para las fórmulas, también su resultado actual (si se conoce).
          const val = sh.values && sh.values[r] ? sh.values[r][c] : undefined;
          if (typeof v === "string" && v.startsWith("=") && val !== undefined && val !== "" && val !== null) {
            s += ` → ${typeof val === "number" ? +val.toPrecision(10) : String(val).replace(/\s+/g, " ")}`;
          }
          if (s.length > o.maxCell) s = s.slice(0, o.maxCell) + "…";
          lines.push(`${cellAddress(sh.rowIndex + r, sh.columnIndex + c)}: ${s}${mark ? " [RELLENAR]" : ""}`);
        }
      }
      if (truncated) lines.push("(hoja recortada)");
      let block = lines.join("\n");
      if (total + block.length > o.maxChars) {
        const room = o.maxChars - total;
        if (room < 200) break;
        block = block.slice(0, room) + "\n(recortado)";
      }
      blocks.push(block);
      total += block.length;
    }
    return blocks.join("\n\n");
  }

  return {
    normalize,
    isEmpty,
    looksLikeQuestion,
    collectQuestions,
    inferExpectedType,
    typeInstruction,
    parseNumber,
    parseDate,
    matchOption,
    coerceAnswer,
    cleanFormula,
    isBrokenFormulaError,
    pickSolveTargets,
    safeText,
    findPastedQuestions,
    ANSWER_SCHEMA,
    SYSTEM_PROMPT,
    buildUserMessage,
    columnLetter,
    cellAddress,
    collectTexts,
    buildWorkbookContext,
  };
});
