/*
 * Respuestas IA — código que corre en segundo plano dentro de Excel.
 *
 * Flujo:
 *  1. Al seleccionar celdas con preguntas (paso previo a copiarlas), se guardan
 *     esos textos solo en memoria: es "la última pregunta copiada".
 *  2. Al seleccionar otra zona (el destino), se toma una foto local de esas celdas:
 *     contenido previo, formato numérico y lista desplegable. Nada de esto sale del equipo.
 *  3. Cuando Excel pega en el destino, cada celda que antes estaba vacía y ahora
 *     contiene una de esas preguntas se envía a la IA (solo el texto de la pregunta
 *     y el tipo de respuesta esperado) y la respuesta sustituye a la pregunta.
 *  4. Si algo falla, la celda no se toca y se añade una nota discreta. Pegar de nuevo reintenta.
 */
import Anthropic from "https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk@1.5.0/+esm";

const L = window.RespuestasLogic;

export const MODEL = "claude-opus-5-5";
const KEY_STORAGE = "respuestas-ia.apiKey";
const MAX_CELLS = 400; // límite por selección/pegado para no bloquear Excel
const BATCH_SIZE = 25; // preguntas por llamada a la IA
const NOTE_MARK = "Respuestas IA:";
const NOTE_TEXT = `${NOTE_MARK} no se pudo obtener la respuesta. Vuelve a pegar la pregunta aquí para reintentar.`;
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
    timeout: 60_000,
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
  cache: new Map(), // `${tipo}|${pregunta}` -> respuesta de la IA
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

    const questions = L.collectQuestions(sel.values);
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

  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheetId);
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
  });

  if (items.length === 0) return;
  items.forEach((it) => state.pending.add(it.key));

  // La IA trabaja fuera de la cola para no bloquear Excel mientras responde.
  answerAndWrite(sheetId, items).finally(() => items.forEach((it) => state.pending.delete(it.key)));
}

async function answerAndWrite(sheetId, items) {
  const apiKey = readKey();
  if (!apiKey) {
    items.forEach((it) => (it.result = { ok: false, reason: "no_key" }));
  } else {
    await resolveAnswers(apiKey, items);
  }
  await enqueue(() => writeResults(sheetId, items));
}

async function resolveAnswers(apiKey, items) {
  const toAsk = [];
  for (const it of items) {
    const cached = state.cache.get(`${it.expected.kind}|${it.question}`);
    if (cached) it.ai = cached;
    else toAsk.push(it);
  }

  if (toAsk.length > 0) {
    const client = makeClient(apiKey);
    // Preguntas repetidas en el mismo pegado se preguntan una sola vez.
    const unique = new Map();
    for (const it of toAsk) {
      const k = `${it.expected.kind}|${it.question}`;
      if (!unique.has(k)) unique.set(k, { id: unique.size, question: it.question, kind: it.expected.kind, k });
    }
    const list = [...unique.values()];
    for (let i = 0; i < list.length; i += BATCH_SIZE) {
      const batch = list.slice(i, i + BATCH_SIZE);
      try {
        const answers = await askAI(client, batch);
        const byId = new Map(answers.map((a) => [a.id, a]));
        for (const q of batch) {
          const a = byId.get(q.id);
          if (a && a.status === "ok") state.cache.set(q.k, a);
          q.ai = a || null;
        }
      } catch (err) {
        console.warn("[Respuestas IA] Error al consultar la IA:", err);
        batch.forEach((q) => (q.ai = null));
      }
    }
    for (const it of toAsk) it.ai = unique.get(`${it.expected.kind}|${it.question}`).ai;
  }

  for (const it of items) {
    it.result = it.ai ? L.coerceAnswer(it.ai, it.expected) : { ok: false, reason: "error" };
  }
}

async function askAI(client, batch) {
  const response = await client.beta.messages.create({
    model: MODEL,
    max_tokens: 16000,
    betas: ["server-side-fallback-2026-07-01"],
    fallbacks: "default", // si el modelo principal rechaza la petición, el servidor prueba con otro
    output_config: {
      effort: "medium",
      format: { type: "json_schema", schema: L.ANSWER_SCHEMA },
    },
    system: L.SYSTEM_PROMPT,
    messages: [{ role: "user", content: L.buildUserMessage(batch) }],
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
      it.address = cell.address;
      // Si el usuario cambió la celda mientras la IA respondía, no se toca.
      if (cell.values[0][0] !== it.question) {
        state.failed.delete(it.key);
        return;
      }
      if (it.result.ok) {
        if (it.snap.numberFormat != null) cell.numberFormat = [[it.snap.numberFormat]];
        cell.values = [[it.result.value]];
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
