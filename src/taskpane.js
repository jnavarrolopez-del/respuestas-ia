/*
 * Panel «Análisis de reservas» (Cuotivo, versión de demostración).
 * - Cargar taller de demostración: hoja nueva con reservas inventadas.
 * - Analizar reservas: cifras calculadas aquí con código; recomendaciones de Claude a través
 *   del servidor (/api/recomendaciones), que guarda la clave. A ese servidor solo van cifras agregadas.
 * - Crear gráfico: hoja nueva con las reservas por día y un gráfico de columnas.
 */
const R = window.Reservas;
const API_RECOMENDACIONES = "/api/recomendaciones";
const MAX_FILAS = 20000;

const $ = (id) => document.getElementById(id);
const botones = [$("btnDemo"), $("btnAnalizar"), $("btnGrafico")];
let ultimaHojaDemo = null;
let ocupado = false;

function estado(texto, tipo) {
  $("estado").textContent = texto;
  $("estado").className = tipo || "";
}

function bloquear(si) {
  ocupado = si;
  botones.forEach((b) => (b.disabled = si));
}

async function tarea(fn) {
  if (ocupado) return;
  bloquear(true);
  try {
    await fn();
  } catch (err) {
    console.error("[Análisis de reservas]", err);
    estado("Algo ha fallado en Excel. Vuelve a intentarlo.", "error");
  } finally {
    bloquear(false);
  }
}

/** Nombre de hoja libre: «Base», «Base 2», «Base 3»… */
async function nombreLibre(ctx, base) {
  const hojas = ctx.workbook.worksheets;
  hojas.load("items/name");
  await ctx.sync();
  const usados = new Set(hojas.items.map((h) => h.name.toLowerCase()));
  if (!usados.has(base.toLowerCase())) return base;
  for (let i = 2; ; i++) if (!usados.has(`${base} ${i}`.toLowerCase())) return `${base} ${i}`;
}

/* ---------- 1. Cargar taller de demostración ---------- */

async function cargarDemo() {
  estado("Creando las reservas del taller de demostración…");
  const reservas = R.generarReservasDemo({ semilla: Date.now() % 1_000_000_007 });
  const filas = R.aFilasExcel(reservas);
  await Excel.run(async (ctx) => {
    const nombre = await nombreLibre(ctx, "Reservas demo");
    const hoja = ctx.workbook.worksheets.add(nombre);
    const rango = hoja.getRangeByIndexes(0, 0, filas.length, filas[0].length);
    rango.values = filas;
    hoja.getRangeByIndexes(1, 0, filas.length - 1, 1).numberFormat = filas.slice(1).map(() => ["dd/mm/yyyy"]);
    hoja.getRangeByIndexes(1, 1, filas.length - 1, 1).numberFormat = filas.slice(1).map(() => ["hh:mm"]);
    const tabla = hoja.tables.add(rango, true);
    tabla.style = "TableStyleMedium7";
    hoja.freezePanes.freezeRows(1);
    rango.format.autofitColumns();
    hoja.activate();
    await ctx.sync();
    ultimaHojaDemo = nombre;
  });
  estado(`Listo: ${reservas.length} reservas en la hoja «${ultimaHojaDemo}».`, "ok");
}

/* ---------- Buscar la hoja de reservas ---------- */

async function leerHoja(ctx, hoja) {
  const usado = hoja.getUsedRangeOrNullObject(true);
  usado.load("isNullObject,rowCount");
  await ctx.sync();
  if (usado.isNullObject || usado.rowCount > MAX_FILAS) return null;
  usado.load("values");
  await ctx.sync();
  return R.leerFilasExcel(usado.values);
}

/** Reservas de la hoja activa; si no es de reservas, de la última de demostración o de otra «Reservas…». */
async function obtenerReservas() {
  return Excel.run(async (ctx) => {
    const activa = ctx.workbook.worksheets.getActiveWorksheet();
    activa.load("name");
    await ctx.sync();
    let leido = await leerHoja(ctx, activa);
    if (leido && leido.reservas.length) return { hoja: activa.name, ...leido };

    const hojas = ctx.workbook.worksheets;
    hojas.load("items/name");
    await ctx.sync();
    const candidatas = hojas.items
      .map((h) => h.name)
      .filter((n) => n !== activa.name && (n === ultimaHojaDemo || /^reservas/i.test(n)))
      .sort((a, b) => (a === ultimaHojaDemo ? -1 : b === ultimaHojaDemo ? 1 : b.localeCompare(a)));
    for (const nombre of candidatas) {
      leido = await leerHoja(ctx, ctx.workbook.worksheets.getItem(nombre));
      if (leido && leido.reservas.length) return { hoja: nombre, ...leido };
    }
    return null;
  });
}

const SIN_RESERVAS =
  "No encuentro reservas. Pulsa «Cargar taller de demostración» o abre una hoja con las columnas Fecha, Hora de entrega, Servicio y Estado.";

/* ---------- 2. Analizar reservas ---------- */

function fechaCorta(iso) {
  if (!iso) return "–";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

const num = (n) => n.toLocaleString("es-ES", { maximumFractionDigits: 1 });

function pintarBarras(ul, filas, max) {
  ul.replaceChildren(
    ...filas.map(({ etiqueta, valor, texto }) => {
      const li = document.createElement("li");
      const nombre = document.createElement("span");
      nombre.textContent = etiqueta;
      const barra = document.createElement("span");
      barra.className = "barra";
      barra.setAttribute("aria-hidden", "true");
      const relleno = document.createElement("i");
      relleno.style.width = `${max ? Math.round((valor / max) * 100) : 0}%`;
      barra.append(relleno);
      const cifra = document.createElement("span");
      cifra.className = "valor";
      cifra.textContent = texto;
      li.append(nombre, barra, cifra);
      return li;
    })
  );
}

function mostrarAnalisis(a, hoja, ignoradas) {
  $("rTotal").textContent = num(a.total);
  $("rMedia").textContent = num(a.mediaPorDia);
  $("rAceptadas").textContent = `${num(a.estados.aceptadas)} %`;
  $("rPeriodo").textContent =
    `Hoja «${hoja}», del ${fechaCorta(a.desde)} al ${fechaCorta(a.hasta)}, ${a.diasConReservas} días con reservas.` +
    (ignoradas ? ` Se han ignorado ${ignoradas} filas incompletas.` : "");

  const maxDia = Math.max(...a.diasSemana.map((d) => d.reservas), 0);
  pintarBarras($("rDias"), a.diasSemana.map((d) => ({ etiqueta: d.dia, valor: d.reservas, texto: String(d.reservas) })), maxDia);

  const maxFranja = Math.max(...a.franjasMasLlenas.map((f) => f.reservas), 0);
  const franjas = (lista) => lista.map((f) => ({ etiqueta: f.franja, valor: f.reservas, texto: String(f.reservas) }));
  pintarBarras($("rLlenas"), franjas(a.franjasMasLlenas), maxFranja);
  pintarBarras($("rVacias"), franjas(a.franjasMasVacias), maxFranja);

  const top = a.servicios.slice(0, 5);
  pintarBarras(
    $("rServicios"),
    top.map((s) => ({ etiqueta: s.servicio, valor: s.reservas, texto: `${s.reservas} · ${num(s.porcentaje)} %` })),
    Math.max(...top.map((s) => s.reservas), 0)
  );

  const e = a.estados;
  pintarBarras(
    $("rEstados"),
    [
      ["Aceptadas", e.aceptadas],
      ["Reprogramadas", e.reprogramadas],
      ["Canceladas", e.canceladas],
      ["Rechazadas", e.rechazadas],
      ["No presentadas", e.no_presentadas],
      ["Pendientes", e.pendientes],
    ].map(([etiqueta, p]) => ({ etiqueta, valor: p, texto: `${num(p)} %` })),
    100
  );
  $("resultados").hidden = false;
}

async function pedirRecomendaciones(agregados) {
  const lista = $("rRecs");
  const aviso = $("rRecsEstado");
  lista.replaceChildren();
  aviso.textContent = "Pidiendo recomendaciones a Claude…";
  const control = new AbortController();
  const plazo = setTimeout(() => control.abort(), 90_000);
  try {
    const resp = await fetch(API_RECOMENDACIONES, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(agregados),
      signal: control.signal,
    });
    let datos = null;
    try {
      datos = await resp.json();
    } catch {
      // Respuesta que no es JSON: no hay servidor de recomendaciones en esta dirección.
    }
    if (!resp.ok || !datos || !Array.isArray(datos.recomendaciones)) {
      aviso.textContent =
        (datos && datos.error) ||
        "El servidor de recomendaciones no está disponible aquí. Las cifras de arriba sí son correctas.";
      return;
    }
    aviso.textContent = "";
    lista.replaceChildren(
      ...datos.recomendaciones.map((t) => {
        const li = document.createElement("li");
        li.textContent = t;
        return li;
      })
    );
  } catch {
    aviso.textContent = "No se ha podido conectar con el servidor de recomendaciones. Las cifras de arriba sí son correctas.";
  } finally {
    clearTimeout(plazo);
  }
}

async function analizar() {
  estado("Leyendo las reservas…");
  const datos = await obtenerReservas();
  if (!datos) {
    estado(SIN_RESERVAS, "error");
    return;
  }
  const analisis = R.analizarReservas(datos.reservas);
  mostrarAnalisis(analisis, datos.hoja, datos.ignoradas);
  estado(`Análisis listo: ${analisis.total} reservas.`, "ok");
  $("resultados").scrollIntoView({ behavior: "smooth", block: "start" });
  // Solo cifras agregadas: nunca matrículas ni reservas sueltas.
  await pedirRecomendaciones(R.agregadosParaIA(analisis));
}

/* ---------- 3. Crear gráfico ---------- */

async function crearGrafico() {
  estado("Preparando el gráfico de reservas por día…");
  const datos = await obtenerReservas();
  if (!datos) {
    estado(SIN_RESERVAS, "error");
    return;
  }
  const porDia = R.reservasPorDia(datos.reservas);
  const filas = [["Fecha", "Reservas"], ...porDia.map((d) => [R.fechaASerial(d.fecha), d.reservas])];
  let nombre;
  await Excel.run(async (ctx) => {
    nombre = await nombreLibre(ctx, "Gráfico de reservas");
    const hoja = ctx.workbook.worksheets.add(nombre);
    const rango = hoja.getRangeByIndexes(0, 0, filas.length, 2);
    rango.values = filas;
    hoja.getRangeByIndexes(1, 0, filas.length - 1, 1).numberFormat = porDia.map(() => ["dd/mm"]);
    const cab = hoja.getRange("A1:B1");
    cab.format.font.bold = true;
    cab.format.fill.color = "#107C41";
    cab.format.font.color = "#FFFFFF";
    rango.format.autofitColumns();

    const grafico = hoja.charts.add(Excel.ChartType.columnClustered, rango, Excel.ChartSeriesBy.columns);
    grafico.title.text = `Reservas por día · ${datos.hoja}`;
    grafico.legend.visible = false;
    grafico.axes.categoryAxis.title.text = "Día";
    grafico.axes.valueAxis.title.text = "Reservas";
    grafico.series.getItemAt(0).format.fill.setSolidColor("#107C41");
    grafico.setPosition("D2", "R26");
    hoja.activate();
    await ctx.sync();
  });
  estado(`Gráfico creado en la hoja «${nombre}» (${porDia.length} días).`, "ok");
}

/* ---------- Arranque ---------- */

$("btnDemo").addEventListener("click", () => tarea(cargarDemo));
$("btnAnalizar").addEventListener("click", () => tarea(analizar));
$("btnGrafico").addEventListener("click", () => tarea(crearGrafico));

Office.onReady(async (info) => {
  // Limpieza de la versión anterior: la clave de Claude ya no se guarda en el navegador
  // y el complemento ya no arranca solo en segundo plano.
  try {
    localStorage.removeItem("respuestas-ia.apiKey");
  } catch {
    // Sin acceso al almacenamiento: no hay nada que borrar.
  }
  if (info.host !== Office.HostType.Excel) {
    estado("Este panel funciona dentro de Excel.", "error");
    return;
  }
  try {
    await Office.addin.setStartupBehavior(Office.StartupBehavior.none);
  } catch {
    // Excel sin runtime compartido: no hace falta.
  }
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
    estado("Esta versión de Excel es demasiado antigua para el análisis de reservas.", "error");
    return;
  }
  bloquear(false);
  estado("Elige qué quieres hacer con las reservas del taller.");
});
