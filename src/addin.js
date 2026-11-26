/*
 * Respuestas IA — código que corre en segundo plano dentro de Excel.
 *
 * Flujo:
 *  1. Al seleccionar celdas con texto (paso previo a copiarlas), se guardan esos textos
 *     solo en memoria: es "lo último copiado" (una pregunta, una etiqueta o un encabezado).
 *  2. Al seleccionar otra zona (el destino), se toma una foto local de esas celdas:
 *     contenido previo, formato numérico y lista desplegable.
 *  3. Cuando Excel pega en el destino, cada celda que antes estaba vacía y ahora contiene
 *     uno de esos textos se envía a la IA junto con el contenido del libro (valores y
 *     fórmulas de sus hojas). La IA devuelve una fórmula (si se puede calcular con el libro)
 *     o un valor, que sustituye al texto pegado.
 *  4. Si algo falla, o la fórmula da error, la celda queda con lo pegado y se añade una nota.
 *     Pegar de nuevo reintenta.
 */
import Anthropic from "https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk@1.5.0/+esm";

const L = window.RespuestasLogic;

export const MODEL = "claude-opus-5-5";
const KEY_STORAGE = "respuestas-ia.apiKey";
const MAX_CELLS = 400; // límite por selección/pegado para no bloquear Excel
const BATCH_SIZE = 25; // celdas por llamada a la IA
const CONTEXT_MAX_SHEETS = 20; // hojas que se leen como contexto
const CONTEXT_MAX_ROWS = 120;
const CONTEXT_MAX_COLS = 26;
const NOTE_MARK = "Respuestas IA:";
const NOTE_TEXT = `${NOTE_MARK} no se pudo obtener la respuesta. Vuelve a pegar la pregunta aquí para reintentar.`;
const NOTE_FORMULA = `${NOTE_MARK} la fórmula propuesta daba error y no se ha escrito. Vuelve a pegar aquí para reintentar.`;
const NOTE_NO_KEY = `${NOTE_MARK} falta la clave de acceso a la IA (Inicio > Configurar clave). Después, vuelve a pegar.`;

/* ---------- Clave de acceso ---------- */

export function readKey() {
  try {
    return localStorage.getItem(KEY_STORAGE) || "";
  } catch {
    return "";
  }
}

export function writeKey(key) {
  localStorage.setItem(KEY_STORAGE, key);
}

export function makeClient(apiKey) {
  return new Anthropic({
    apiKey,
    dangerouslyAllowBrowser: true, // complemento personal: la clave es del propio usuario y vive solo en este equipo
    timeout: 180_000, // con el contexto del libro, la IA puede tardar más
    maxRetries: 2,
  });
}

/* ---------- Estado en memoria (se pierde al cerrar el libro) ---------- */

const state = {
  sourceQuestions: new Set(),
  sourceDims: { rows: 1, cols: 1 },
  snapshots: new Map(), // clave de celda -> { value, numberFormat, listOptions, listRule, header }
  failed: new Map(), // clave de celda -> { question, snap, marked }
  pending: new Set(),
};

const cellKey = (sheetId, row, col) => `${sheetId}|${row}|${col}`;

// Las lecturas de Excel se hacen en orden para que la foto del destino exista antes de procesar el pegado.
let excelQueue = Promise.resolve();
function enqueue(fn) {
  const run = excelQueue.then(fn, fn);
  excelQueue = run.catch(() => {});
  return run;
}

function quiet(fn) {
  return (e) =>
    enqueue(() => fn(e)).catch((err) => {
      console.warn("[Respuestas IA]", err);
    });
}

/* ---------- Registro de eventos ---------- */

let registered = false;

export async function registerHandlers() {
  if (registered) return;
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    console.warn("[Respuestas IA] Esta versión de Excel no admite los eventos necesarios (ExcelApi 1.9).");
    return;
  }
  registered = true;
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.onSelectionChanged.add(quiet(onSelectionChanged));
    ctx.workbook.worksheets.onChanged.add(quiet(onChanged));
    await ctx.sync();
  });
}

/* ---------- Selección: origen (preguntas) o destino (foto previa) ---------- */

async function onSelectionChanged(e) {
  if (!e.address || e.address.includes(",")) return; // selecciones múltiples: se ignoran
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(e.worksheetId);
    const sel = ws.getRange(e.address);
    sel.load("rowIndex,columnIndex,rowCount,columnCount,cellCount");
    await ctx.sync();
    if (sel.cellCount > MAX_CELLS) return;

    sel.load("values");
    await ctx.sync();

    const questions = L.collectTexts(sel.values);
    if (questions.length > 0) {
      state.sourceQuestions = new Set(questions);
      state.sourceDims = { rows: sel.rowCount, cols: sel.columnCount };
      return;
    }
    if (state.sourceQuestions.size === 0) return;

    // Destino probable: fotografía del área donde caerá el pegado.
    const rows = Math.max(sel.rowCount, state.sourceDims.rows);
    const cols = Math.max(sel.columnCount, state.sourceDims.cols);
    if (rows * cols > MAX_CELLS) return;
    await snapshotRegion(ctx, ws, e.worksheetId, sel.rowIndex, sel.columnIndex, rows, cols);
  });
}

async function snapshotRegion(ctx, ws, sheetId, row0, col0, rows, cols) {
  const region = ws.getRangeByIndexes(row0, col0, rows, cols);
  region.load("values,numberFormat");
  let header = null;
  if (row0 > 0) {
    header = ws.getRangeByIndexes(0, col0, 1, cols);
    header.load("values");
  }
  const dvs = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const dv = region.getCell(r, c).dataValidation;
      dv.load("type,rule");
      dvs.push(dv);
    }
  }
  await ctx.sync();

  const listSources = new Map(); // fuente de la lista -> opciones
  for (const dv of dvs) {
    if (dv.type === "List" && dv.rule && dv.rule.list && dv.rule.list.source != null) {
      listSources.set(String(dv.rule.list.source), null);
    }
  }
  await resolveListSources(ctx, ws, listSources);

  let i = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const dv = dvs[i++];
      const isList = dv.type === "List" && dv.rule && dv.rule.list;
      const headerText = header ? header.values[0][c] : null;
      state.snapshots.set(cellKey(sheetId, row0 + r, col0 + c), {
        value: region.values[r][c],
        numberFormat: region.numberFormat[r][c],
        listOptions: isList ? listSources.get(String(dv.rule.list.source)) || null : null,
        listRule: isList
          ? { list: { inCellDropDown: dv.rule.list.inCellDropDown !== false, source: dv.rule.list.source } }
          : null,
        header: typeof headerText === "string" && !L.looksLikeQuestion(headerText) ? headerText : null,
      });
    }
  }
  // Mantener la memoria acotada.
  if (state.snapshots.size > 5000) {
    const extra = state.snapshots.size - 5000;
    let n = 0;
    for (const k of state.snapshots.keys()) {
      if (n++ >= extra) break;
      state.snapshots.delete(k);
    }
  }
}

/** Convierte las fuentes de listas desplegables ("Sí,No" o "=$D$1:$D$4") en opciones. */
async function resolveListSources(ctx, ws, listSources) {
  for (const src of [...listSources.keys()]) {
    const s = src.trim();
    if (!s.startsWith("=")) {
      listSources.set(src, s.split(/[,;]/).map((x) => x.trim()).filter(Boolean));
      continue;
    }
    // Cada referencia se resuelve por separado: si una falla, no arrastra a las demás.
    try {
      const ref = s.slice(1);
      let range;
      const bang = ref.lastIndexOf("!");
      if (bang > -1) {
        const sheetName = ref.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'");
        range = ctx.workbook.worksheets.getItem(sheetName).getRange(ref.slice(bang + 1));
      } else if (/^\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?$/i.test(ref)) {
        range = ws.getRange(ref);
      } else {
        range = ctx.workbook.names.getItem(ref).getRange();
      }
      range.load("values");
      await ctx.sync();
      listSources.set(src, range.values.flat().map((v) => String(v).trim()).filter(Boolean));
    } catch {
      listSources.set(src, null);
    }
  }
}

/* ---------- Pegado: detectar preguntas pegadas y responder ---------- */

async function onChanged(e) {
  if (e.source && e.source !== "Local") return;
  if (e.changeType && e.changeType !== "RangeEdited") return;
  if (state.sourceQuestions.size === 0) return;
  if (!e.address || e.address.includes(",")) return;

  const sheetId = e.worksheetId;
  const sourceSet = new Set(state.sourceQuestions);
  let items = [];
  let context = "";

  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheetId);
    ws.load("name");
    const rng = ws.getRange(e.address);
    rng.load("rowIndex,columnIndex,rowCount,columnCount,cellCount");
    await ctx.sync();
    if (rng.cellCount > MAX_CELLS) return;
    rng.load("values");
    await ctx.sync();

    const single = rng.cellCount === 1;
    const keyAt = (r, c) => cellKey(sheetId, rng.rowIndex + r, rng.columnIndex + c);

    const found = L.findPastedQuestions(
      rng.values,
      (r, c) => {
        const snap = state.snapshots.get(keyAt(r, c));
        if (snap) return { known: true, value: snap.value };
        if (single && e.details) return { known: true, value: e.details.valueBefore };
        return { known: false };
      },
      sourceSet,
      (r, c, v) => {
        const f = state.failed.get(keyAt(r, c));
        return !!f && f.question === v;
      }
    );

    for (const f of found) {
      const key = keyAt(f.r, f.c);
      if (state.pending.has(key)) continue;
      const prevFail = state.failed.get(key);
      const snap =
        (prevFail && prevFail.question === f.question && prevFail.snap) ||
        state.snapshots.get(key) || { value: "", numberFormat: null, listOptions: null, listRule: null, header: null };
      items.push({
        key,
        id: items.length,
        sheet: ws.name,
        address: L.cellAddress(rng.rowIndex + f.r, rng.columnIndex + f.c),
        row: rng.rowIndex + f.r,
        col: rng.columnIndex + f.c,
        question: f.question,
        snap,
        expected: L.inferExpectedType(snap),
        hadMarker: !!(prevFail && prevFail.marked),
      });
    }

    // La foto de las celdas cambiadas ya no es válida.
    for (let r = 0; r < rng.rowCount; r++) {
      for (let c = 0; c < rng.columnCount; c++) state.snapshots.delete(keyAt(r, c));
    }

    if (items.length > 0) {
      try {
        context = await readWorkbookContext(ctx, ws.name);
      } catch (err) {
        console.warn("[Respuestas IA] No se pudo leer el libro como contexto:", err);
      }
    }
  });

  if (items.length === 0) return;
  items.forEach((it) => state.pending.add(it.key));

  // La IA trabaja fuera de la cola para no bloquear Excel mientras responde.
  answerAndWrite(sheetId, items, context).finally(() => items.forEach((it) => state.pending.delete(it.key)));
}

/** Lee todas las hojas (valores, fórmulas y celdas vacías resaltadas para rellenar). */
async function readWorkbookContext(ctx, activeName) {
  const sheets = ctx.workbook.worksheets;
  sheets.load("items/name,items/visibility");
  await ctx.sync();
  const visible = sheets.items.filter((s) => s.visibility === "Visible").slice(0, CONTEXT_MAX_SHEETS);
  const used = visible.map((s) => {
    const u = s.getUsedRangeOrNullObject(false);
    u.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
    return { s, u };
  });
  await ctx.sync();

  const parts = [];
  for (const { s, u } of used) {
    if (u.isNullObject) continue;
    const rows = Math.min(u.rowCount, CONTEXT_MAX_ROWS);
    const cols = Math.min(u.columnCount, CONTEXT_MAX_COLS);
    const r = s.getRangeByIndexes(u.rowIndex, u.columnIndex, rows, cols);
    r.load("formulas,values");
    let props = null;
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
      props = r.getCellProperties({ format: { fill: { color: true } } });
    }
    parts.push({ name: s.name, u, r, props, rows, cols });
  }
  await ctx.sync();

  const data = parts.map((p) => ({
    name: p.name,
    rowIndex: p.u.rowIndex,
    columnIndex: p.u.columnIndex,
    formulas: p.r.formulas,
    values: p.r.values,
    marks: p.props
      ? p.props.value.map((row, i) =>
          row.map((cp, j) => L.isEmpty(p.r.formulas[i][j]) && isLightFill(cp && cp.format && cp.format.fill && cp.format.fill.color))
        )
      : null,
  }));
  return L.buildWorkbookContext(data, activeName, { maxRows: CONTEXT_MAX_ROWS, maxCols: CONTEXT_MAX_COLS });
}

/** Relleno claro y de color (no blanco): así suelen marcarse las celdas a completar. */
function isLightFill(color) {
  if (typeof color !== "string" || !/^#[0-9a-f]{6}$/i.test(color)) return false;
  const n = parseInt(color.slice(1), 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  if (r > 245 && g > 245 && b > 245) return false; // blanco
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 200;
}

async function answerAndWrite(sheetId, items, context) {
  const apiKey = readKey();
  if (!apiKey) {
    items.forEach((it) => (it.result = { ok: false, reason: "no_key" }));
  } else {
    await resolveAnswers(apiKey, items, context);
  }
  await enqueue(() => writeResults(sheetId, items));
}

async function resolveAnswers(apiKey, items, context) {
  const client = makeClient(apiKey);
  for (let i = 0; i < items.length; i += BATCH_SIZE) {
    const batch = items.slice(i, i + BATCH_SIZE);
    try {
      const answers = await askAI(
        client,
        batch.map((it) => ({ id: it.id, sheet: it.sheet, address: it.address, question: it.question, kind: it.expected.kind })),
        context
      );
      const byId = new Map(answers.map((a) => [a.id, a]));
      batch.forEach((it) => (it.ai = byId.get(it.id) || null));
    } catch (err) {
      console.warn("[Respuestas IA] Error al consultar la IA:", err);
      batch.forEach((it) => (it.ai = null));
    }
  }
  for (const it of items) {
    it.result = it.ai ? L.coerceAnswer(it.ai, it.expected) : { ok: false, reason: "error" };
  }
}

async function askAI(client, batch, context) {
  const response = await client.beta.messages.create({
    model: MODEL,
    max_tokens: 16000,
    betas: ["server-side-fallback-2026-07-01"],
    fallbacks: "default", // si el modelo principal rechaza la petición, el servidor prueba con otro
    output_config: {
      effort: "high", // escribir fórmulas correctas a partir del libro merece pensar más
      format: { type: "json_schema", schema: L.ANSWER_SCHEMA },
    },
    system: L.SYSTEM_PROMPT,
    messages: [{ role: "user", content: L.buildUserMessage(batch, context) }],
  });
  if (response.stop_reason === "refusal" || response.stop_reason === "max_tokens") {
    throw new Error(`Respuesta incompleta (${response.stop_reason})`);
  }
  const text = response.content
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("");
  const parsed = JSON.parse(text);
  return Array.isArray(parsed.answers) ? parsed.answers : [];
}

async function writeResults(sheetId, items) {
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(sheetId);
    ws.load("isNullObject");
    await ctx.sync();
    if (ws.isNullObject) return;

    const cells = items.map((it) => {
      const cell = ws.getRangeByIndexes(it.row, it.col, 1, 1);
      cell.load("values,address");
      cell.dataValidation.load("type");
      return cell;
    });
    await ctx.sync();

    const toMark = [];
    const toUnmark = [];
    items.forEach((it, i) => {
      const cell = cells[i];
      // Si el usuario cambió la celda mientras la IA respondía, no se toca.
      if (cell.values[0][0] !== it.question) {
        it.skipped = true;
        state.failed.delete(it.key);
        return;
      }
      if (it.result.ok) {
        if (it.snap.numberFormat != null) cell.numberFormat = [[it.snap.numberFormat]];
        if (it.result.formula) {
          cell.formulas = [[it.result.formula]];
          cell.load("valueTypes");
          it.checkFormula = cell;
        } else {
          cell.values = [[it.result.value]];
        }
        it.restoreList = it.snap.listRule && cell.dataValidation.type !== "List" ? cell : null;
        if (it.hadMarker) toUnmark.push(cell);
        state.failed.delete(it.key);
      } else if (it.result.reason === "not_a_question") {
        // Estado vacío: no era una pregunta; se deja la celda como está.
        state.failed.delete(it.key);
      } else {
        const marked = it.hadMarker;
        state.failed.set(it.key, { question: it.question, snap: it.snap, marked: true });
        if (!marked) toMark.push({ cell, text: it.result.reason === "no_key" ? NOTE_NO_KEY : NOTE_TEXT });
      }
    });
    try {
      await ctx.sync();
    } catch (err) {
      // Excel rechazó alguna fórmula (sintaxis no válida): se vuelve a lo pegado celda a celda.
      console.warn("[Respuestas IA] Excel rechazó la escritura; se reintenta celda a celda:", err);
      await writeOneByOne(ctx, ws, items, toMark, toUnmark);
    }

    // Una fórmula que da error (#¿NOMBRE?, #REF!, #DIV/0!…) no cuenta como respuesta.
    for (const it of items) {
      if (!it.checkFormula) continue;
      let types;
      try {
        types = it.checkFormula.valueTypes;
      } catch {
        continue;
      }
      if (types && types[0] && types[0][0] === "Error") {
        it.checkFormula.values = [[it.question]];
        state.failed.set(it.key, { question: it.question, snap: it.snap, marked: true });
        const i = toUnmark.indexOf(it.checkFormula);
        if (i > -1) toUnmark.splice(i, 1);
        else if (!it.hadMarker) toMark.push({ cell: it.checkFormula, text: NOTE_FORMULA });
      }
    }
    await ctx.sync();

    // Pasos secundarios: si fallan, la respuesta ya está escrita.
    for (const it of items) {
      if (!it.restoreList) continue;
      try {
        it.restoreList.dataValidation.clear();
        it.restoreList.dataValidation.rule = it.snap.listRule;
        await ctx.sync();
      } catch (err) {
        console.warn("[Respuestas IA] No se pudo restaurar la lista desplegable:", err);
      }
    }
    for (const cell of toUnmark) await removeMarker(ctx, ws, cell);
    for (const m of toMark) await addMarker(ctx, ws, m.cell, m.text);
  });
}

/** Plan B si un lote falla: escribe cada fórmula por separado y deja lo pegado donde Excel no la acepte. */
async function writeOneByOne(ctx, ws, items, toMark) {
  for (const it of items) {
    it.checkFormula = null;
    if (!it.result.ok || it.skipped) continue;
    const cell = ws.getRangeByIndexes(it.row, it.col, 1, 1);
    cell.load("address");
    try {
      if (it.snap.numberFormat != null) cell.numberFormat = [[it.snap.numberFormat]];
      if (it.result.formula) {
        cell.formulas = [[it.result.formula]];
        cell.load("valueTypes");
      } else {
        cell.values = [[it.result.value]];
      }
      await ctx.sync();
      if (it.result.formula) it.checkFormula = cell;
    } catch {
      it.checkFormula = null;
      try {
        cell.values = [[it.question]];
        await ctx.sync();
      } catch {
        // Sin más opciones: la celda queda como Excel la dejó.
      }
      state.failed.set(it.key, { question: it.question, snap: it.snap, marked: true });
      if (!it.hadMarker) toMark.push({ cell, text: NOTE_FORMULA });
    }
  }
}

/* ---------- Marca discreta de error (nota; si no hay notas, comentario) ---------- */

async function addMarker(ctx, ws, cell, text) {
  try {
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.18")) {
      ws.notes.add(cell, text);
    } else if (Office.context.requirements.isSetSupported("ExcelApi", "1.10")) {
      ctx.workbook.comments.add(cell, text);
    } else {
      return;
    }
    await ctx.sync();
  } catch (err) {
    console.warn("[Respuestas IA] No se pudo añadir la nota:", err);
  }
}

async function removeMarker(ctx, ws, cell) {
  try {
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.18")) {
      const notes = ws.notes;
      notes.load("items/content");
      await ctx.sync();
      const locs = notes.items.map((n) => {
        const loc = n.getLocation();
        loc.load("address");
        return { n, loc };
      });
      await ctx.sync();
      for (const { n, loc } of locs) {
        if (loc.address === cell.address && String(n.content).startsWith(NOTE_MARK)) n.delete();
      }
      await ctx.sync();
    } else if (Office.context.requirements.isSetSupported("ExcelApi", "1.10")) {
      const comment = ctx.workbook.comments.getItemByCell(cell);
      comment.load("content");
      await ctx.sync();
      if (String(comment.content).startsWith(NOTE_MARK)) {
        comment.delete();
        await ctx.sync();
      }
    }
  } catch {
    // Sin nota que quitar.
  }
}
