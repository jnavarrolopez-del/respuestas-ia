/*
 * Lógica pura del análisis de reservas (sin Office ni red).
 * Se usa en el navegador (window.Reservas) y en las pruebas con Node.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Reservas = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const SERVICIOS = [
    "Aceite y filtros",
    "Revisión",
    "Pre-ITV",
    "Neumáticos",
    "Frenos",
    "Diagnosis",
    "Aire acondicionado",
    "Batería",
    "Otro",
  ];
  const ESTADOS = ["pendiente", "confirmada", "en_taller", "lista", "entregada", "rechazada", "cancelada", "no_presentado"];
  const ACEPTADAS = new Set(["confirmada", "en_taller", "lista", "entregada"]);
  const DIAS = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  const COLUMNAS = ["Fecha", "Hora de entrega", "Matrícula", "Servicio", "Estado", "Reprogramada"];

  /** Franjas de 30 minutos de la jornada partida: 9:00–14:00 y 16:00–19:00. */
  const FRANJAS = (() => {
    const out = [];
    for (const [desde, hasta] of [[9 * 60, 14 * 60], [16 * 60, 19 * 60]]) {
      for (let m = desde; m < hasta; m += 30) out.push(minutosAHora(m));
    }
    return out;
  })();

  function minutosAHora(m) {
    return `${String(Math.floor(m / 60)).padStart(2, "0")}:${String(m % 60).padStart(2, "0")}`;
  }

  /* ---------- Generador de reservas de demostración ---------- */

  /** Generador pseudoaleatorio con semilla (mulberry32): mismas reservas con la misma semilla. */
  function aleatorio(semilla) {
    let a = semilla >>> 0;
    return () => {
      a = (a + 0x6d2b79f5) >>> 0;
      let t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  function elegir(rnd, pesos) {
    const total = pesos.reduce((s, [, p]) => s + p, 0);
    let x = rnd() * total;
    for (const [valor, p] of pesos) {
      x -= p;
      if (x < 0) return valor;
    }
    return pesos[pesos.length - 1][0];
  }

  function fechaISO(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }

  function diaSemana(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    return new Date(y, m - 1, d).getDay();
  }

  /** Días laborables (lunes a viernes) desde hace 2 meses hasta hoy, ambos incluidos. */
  function diasLaborables(hoy) {
    const fin = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
    const inicio = new Date(fin.getFullYear(), fin.getMonth() - 2, fin.getDate());
    const dias = [];
    for (let d = new Date(inicio); d <= fin; d.setDate(d.getDate() + 1)) {
      const s = d.getDay();
      if (s >= 1 && s <= 5) dias.push(fechaISO(d));
    }
    return dias;
  }

  const PESO_DIA = { 1: 1.3, 2: 1.0, 3: 0.95, 4: 0.9, 5: 1.15 };
  const PESO_FRANJA = {
    "09:00": 2.2, "09:30": 1.8, "10:00": 1.5, "10:30": 1.2, "11:00": 1.0, "11:30": 0.9, "12:00": 0.8,
    "12:30": 0.6, "13:00": 0.4, "13:30": 0.25, "16:00": 1.1, "16:30": 0.9, "17:00": 0.7, "17:30": 0.5,
    "18:00": 0.3, "18:30": 0.15,
  };
  const PESO_SERVICIO = [
    ["Aceite y filtros", 24], ["Revisión", 15], ["Neumáticos", 12], ["Frenos", 10], ["Pre-ITV", 10],
    ["Diagnosis", 10], ["Aire acondicionado", 6], ["Batería", 6], ["Otro", 7],
  ];
  const LETRAS_MATRICULA = "BCDFGHJKLMNPRSTVWXYZ";

  function estadoSegunAntiguedad(rnd, diasAtras) {
    if (diasAtras === 0) {
      return elegir(rnd, [["confirmada", 35], ["en_taller", 25], ["pendiente", 20], ["cancelada", 10], ["rechazada", 10]]);
    }
    if (diasAtras <= 3) {
      return elegir(rnd, [
        ["entregada", 45], ["lista", 15], ["en_taller", 12], ["confirmada", 5],
        ["cancelada", 10], ["rechazada", 6], ["no_presentado", 7],
      ]);
    }
    return elegir(rnd, [["entregada", 75], ["cancelada", 10], ["rechazada", 7], ["no_presentado", 8]]);
  }

  /**
   * Unas `total` reservas inventadas de los últimos 2 meses, ordenadas por fecha y hora.
   * Cada una: { fecha: "AAAA-MM-DD", hora: "HH:MM", matricula, servicio, estado, reprogramada: "sí"|"no" }.
   */
  function generarReservasDemo(opciones) {
    const o = Object.assign({ hoy: new Date(), semilla: 20261126, total: 150 }, opciones || {});
    const rnd = aleatorio(o.semilla);
    const dias = diasLaborables(o.hoy);
    const hoyISO = fechaISO(o.hoy);
    const pesosDia = dias.map((d) => [d, PESO_DIA[diaSemana(d)]]);
    const pesosFranja = FRANJAS.map((f) => [f, PESO_FRANJA[f]]);
    const usadas = new Set();
    const reservas = [];
    for (let i = 0; i < o.total; i++) {
      const fecha = elegir(rnd, pesosDia);
      const hora = elegir(rnd, pesosFranja);
      let matricula;
      do {
        const num = String(Math.floor(rnd() * 10000)).padStart(4, "0");
        let letras = "";
        for (let k = 0; k < 3; k++) letras += LETRAS_MATRICULA[Math.floor(rnd() * LETRAS_MATRICULA.length)];
        matricula = `${num} ${letras}`;
      } while (usadas.has(matricula));
      usadas.add(matricula);
      const diasAtras = Math.round((Date.parse(hoyISO) - Date.parse(fecha)) / 86400000);
      const estado = estadoSegunAntiguedad(rnd, diasAtras);
      const reprogramada = estado !== "rechazada" && rnd() < 0.12 ? "sí" : "no";
      reservas.push({ fecha, hora, matricula, servicio: elegir(rnd, PESO_SERVICIO), estado, reprogramada });
    }
    reservas.sort((a, b) => (a.fecha + a.hora).localeCompare(b.fecha + b.hora));
    return reservas;
  }

  /* ---------- Conversión con las celdas de Excel ---------- */

  const DIA_CERO = Date.UTC(1899, 11, 30);

  function fechaASerial(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    return Math.round((Date.UTC(y, m - 1, d) - DIA_CERO) / 86400000);
  }

  function horaASerial(hhmm) {
    const [h, m] = hhmm.split(":").map(Number);
    return (h * 60 + m) / 1440;
  }

  /** Filas para escribir en Excel: cabecera y valores (fecha y hora como números de Excel). */
  function aFilasExcel(reservas) {
    return [
      COLUMNAS.slice(),
      ...reservas.map((r) => [fechaASerial(r.fecha), horaASerial(r.hora), r.matricula, r.servicio, r.estado, r.reprogramada]),
    ];
  }

  function leerFecha(v) {
    if (typeof v === "number" && v > 0) {
      const d = new Date(DIA_CERO + Math.floor(v) * 86400000);
      return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
    }
    if (typeof v !== "string") return null;
    const t = v.trim();
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
    m = t.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/);
    if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
    return null;
  }

  function leerHora(v) {
    if (typeof v === "number" && v >= 0) {
      const minutos = Math.round((v % 1) * 1440) % 1440;
      return minutosAHora(minutos);
    }
    if (typeof v !== "string") return null;
    const m = v.trim().match(/^(\d{1,2}):(\d{2})/);
    return m ? `${m[1].padStart(2, "0")}:${m[2]}` : null;
  }

  function clave(s) {
    return String(s == null ? "" : s)
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  /**
   * Convierte los valores de una hoja (con cabecera en la primera fila) en reservas.
   * Devuelve { reservas, ignoradas } o null si la hoja no tiene las columnas necesarias.
   */
  function leerFilasExcel(valores) {
    if (!Array.isArray(valores) || valores.length < 2) return null;
    const cab = valores[0].map(clave);
    const col = (...nombres) => cab.findIndex((c) => nombres.some((n) => c === n || c.startsWith(n)));
    const iFecha = col("fecha");
    const iHora = col("hora");
    const iServicio = col("servicio");
    const iEstado = col("estado");
    const iRepro = col("reprogramada");
    const iMatricula = col("matricula");
    if ([iFecha, iHora, iServicio, iEstado].some((i) => i < 0)) return null;
    const reservas = [];
    let ignoradas = 0;
    for (const fila of valores.slice(1)) {
      if (fila.every((v) => v === "" || v == null)) continue;
      const fecha = leerFecha(fila[iFecha]);
      const hora = leerHora(fila[iHora]);
      const estado = clave(fila[iEstado]).replace(/ /g, "_");
      if (!fecha || !hora || !ESTADOS.includes(estado)) {
        ignoradas++;
        continue;
      }
      reservas.push({
        fecha,
        hora,
        matricula: iMatricula >= 0 ? String(fila[iMatricula]) : "",
        servicio: String(fila[iServicio] || "Otro").trim() || "Otro",
        estado,
        reprogramada: iRepro >= 0 && ["si", "sí", "true", "1"].includes(clave(fila[iRepro])) ? "sí" : "no",
      });
    }
    return { reservas, ignoradas };
  }

  /* ---------- Análisis (todo con código) ---------- */

  function porcentaje(parte, total) {
    return total ? Math.round((parte / total) * 1000) / 10 : 0;
  }

  /** Franja de 30 minutos a la que pertenece una hora ("09:40" -> "09:30"). */
  function franjaDe(hora) {
    const [h, m] = hora.split(":").map(Number);
    return minutosAHora(h * 60 + (m >= 30 ? 30 : 0));
  }

  function analizarReservas(reservas) {
    const total = reservas.length;
    const fechas = reservas.map((r) => r.fecha).sort();

    const cuentaDias = new Map();
    for (const r of reservas) {
      const d = DIAS[diaSemana(r.fecha)];
      cuentaDias.set(d, (cuentaDias.get(d) || 0) + 1);
    }
    const ordenSemana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
    const diasSemana = ordenSemana
      .filter((d) => cuentaDias.has(d) || ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"].includes(d))
      .map((dia) => ({ dia, reservas: cuentaDias.get(dia) || 0 }))
      .sort((a, b) => b.reservas - a.reservas || ordenSemana.indexOf(a.dia) - ordenSemana.indexOf(b.dia));

    const cuentaFranjas = new Map(FRANJAS.map((f) => [f, 0]));
    for (const r of reservas) {
      const f = franjaDe(r.hora);
      cuentaFranjas.set(f, (cuentaFranjas.get(f) || 0) + 1);
    }
    const franjas = [...cuentaFranjas.entries()].map(([franja, n]) => ({ franja, reservas: n }));
    const porHora = (a, b) => a.franja.localeCompare(b.franja);
    const franjasMasLlenas = [...franjas].sort((a, b) => b.reservas - a.reservas || porHora(a, b)).slice(0, 3);
    const franjasMasVacias = [...franjas].sort((a, b) => a.reservas - b.reservas || porHora(a, b)).slice(0, 3);

    const cuentaServ = new Map();
    for (const r of reservas) cuentaServ.set(r.servicio, (cuentaServ.get(r.servicio) || 0) + 1);
    const servicios = [...cuentaServ.entries()]
      .map(([servicio, n]) => ({ servicio, reservas: n, porcentaje: porcentaje(n, total) }))
      .sort((a, b) => b.reservas - a.reservas || a.servicio.localeCompare(b.servicio));

    const cuenta = (pred) => reservas.filter(pred).length;
    const estados = {
      aceptadas: porcentaje(cuenta((r) => ACEPTADAS.has(r.estado)), total),
      reprogramadas: porcentaje(cuenta((r) => r.reprogramada === "sí"), total),
      canceladas: porcentaje(cuenta((r) => r.estado === "cancelada"), total),
      rechazadas: porcentaje(cuenta((r) => r.estado === "rechazada"), total),
      no_presentadas: porcentaje(cuenta((r) => r.estado === "no_presentado"), total),
      pendientes: porcentaje(cuenta((r) => r.estado === "pendiente"), total),
    };

    const diasConReservas = new Set(fechas).size;
    return {
      total,
      desde: fechas[0] || null,
      hasta: fechas[fechas.length - 1] || null,
      diasConReservas,
      mediaPorDia: diasConReservas ? Math.round((total / diasConReservas) * 10) / 10 : 0,
      diasSemana,
      franjasMasLlenas,
      franjasMasVacias,
      servicios,
      estados,
    };
  }

  /** Solo cifras agregadas para la IA: nunca matrículas ni reservas sueltas. */
  function agregadosParaIA(a) {
    return {
      total: a.total,
      desde: a.desde,
      hasta: a.hasta,
      mediaPorDia: a.mediaPorDia,
      diasSemana: a.diasSemana.map(({ dia, reservas }) => ({ dia, reservas })),
      franjasMasLlenas: a.franjasMasLlenas.map(({ franja, reservas }) => ({ franja, reservas })),
      franjasMasVacias: a.franjasMasVacias.map(({ franja, reservas }) => ({ franja, reservas })),
      servicios: a.servicios.slice(0, 9).map(({ servicio, reservas, porcentaje }) => ({ servicio, reservas, porcentaje })),
      porcentajes: Object.assign({}, a.estados),
    };
  }

  /** Reservas por día, de la primera a la última fecha, con los días laborables sin reservas a 0. */
  function reservasPorDia(reservas) {
    if (!reservas.length) return [];
    const cuenta = new Map();
    for (const r of reservas) cuenta.set(r.fecha, (cuenta.get(r.fecha) || 0) + 1);
    const fechas = [...cuenta.keys()].sort();
    const [y, m, d] = fechas[0].split("-").map(Number);
    const fin = fechas[fechas.length - 1];
    const out = [];
    for (let f = new Date(y, m - 1, d); fechaISO(f) <= fin; f.setDate(f.getDate() + 1)) {
      const iso = fechaISO(f);
      const s = f.getDay();
      if (cuenta.has(iso) || (s >= 1 && s <= 5)) out.push({ fecha: iso, reservas: cuenta.get(iso) || 0 });
    }
    return out;
  }

  return {
    SERVICIOS,
    ESTADOS,
    FRANJAS,
    COLUMNAS,
    generarReservasDemo,
    diasLaborables,
    fechaASerial,
    horaASerial,
    aFilasExcel,
    leerFilasExcel,
    analizarReservas,
    agregadosParaIA,
    reservasPorDia,
    franjaDe,
  };
});
