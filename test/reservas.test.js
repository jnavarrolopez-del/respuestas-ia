// Pruebas de la lógica de reservas. Ejecutar con:  node --test test/
const test = require("node:test");
const assert = require("node:assert/strict");
const R = require("../src/reservas.js");

const HOY = new Date(2026, 10, 26); // jueves 26/11/2026

test("genera unas 150 reservas válidas de los últimos 2 meses, de lunes a viernes", () => {
  const res = R.generarReservasDemo({ hoy: HOY });
  assert.equal(res.length, 150);
  for (const r of res) {
    assert.ok(r.fecha >= "2026-09-26" && r.fecha <= "2026-11-26", r.fecha);
    const s = new Date(r.fecha + "T12:00:00").getDay();
    assert.ok(s >= 1 && s <= 5, `${r.fecha} es fin de semana`);
    assert.ok(R.FRANJAS.includes(r.hora), r.hora);
    assert.match(r.matricula, /^\d{4} [BCDFGHJKLMNPRSTVWXYZ]{3}$/);
    assert.ok(R.SERVICIOS.includes(r.servicio), r.servicio);
    assert.ok(R.ESTADOS.includes(r.estado), r.estado);
    assert.ok(["sí", "no"].includes(r.reprogramada));
  }
  assert.equal(new Set(res.map((r) => r.matricula)).size, 150, "matrículas repetidas");
  // Ordenadas por fecha y hora
  const claves = res.map((r) => r.fecha + r.hora);
  assert.deepEqual(claves, [...claves].sort());
});

test("las franjas son de 30 minutos en jornada partida (9–14 y 16–19)", () => {
  assert.equal(R.FRANJAS.length, 16);
  assert.equal(R.FRANJAS[0], "09:00");
  assert.equal(R.FRANJAS[9], "13:30");
  assert.equal(R.FRANJAS[10], "16:00");
  assert.equal(R.FRANJAS[15], "18:30");
  assert.ok(!R.FRANJAS.includes("14:00") && !R.FRANJAS.includes("19:00"));
});

test("la misma semilla da las mismas reservas y otra semilla, otras", () => {
  const a = R.generarReservasDemo({ hoy: HOY, semilla: 1 });
  const b = R.generarReservasDemo({ hoy: HOY, semilla: 1 });
  const c = R.generarReservasDemo({ hoy: HOY, semilla: 2 });
  assert.deepEqual(a, b);
  assert.notDeepEqual(a, c);
});

test("ida y vuelta por Excel: fechas y horas como números de Excel", () => {
  assert.equal(R.fechaASerial("2026-11-26"), 46352);
  assert.equal(R.horaASerial("09:30"), 9.5 / 24);
  const res = R.generarReservasDemo({ hoy: HOY, total: 20 });
  const filas = R.aFilasExcel(res);
  assert.deepEqual(filas[0], R.COLUMNAS);
  const leido = R.leerFilasExcel(filas);
  assert.equal(leido.ignoradas, 0);
  assert.deepEqual(leido.reservas, res);
});

test("leerFilasExcel acepta textos, ignora filas malas y rechaza hojas sin columnas", () => {
  const filas = [
    ["Fecha", "Hora de entrega", "Matrícula", "Servicio", "Estado", "Reprogramada"],
    ["03/11/2026", "9:00", "1234 BCD", "Frenos", "Entregada", "Sí"],
    ["2026-11-04", "16:30", "5678 FGH", "Batería", "no presentado", "no"],
    ["mañana", "10:00", "1111 JKL", "Otro", "confirmada", "no"],
    ["", "", "", "", "", ""],
  ];
  const { reservas, ignoradas } = R.leerFilasExcel(filas);
  assert.equal(ignoradas, 1);
  assert.deepEqual(reservas[0], { fecha: "2026-11-03", hora: "09:00", matricula: "1234 BCD", servicio: "Frenos", estado: "entregada", reprogramada: "sí" });
  assert.equal(reservas[1].estado, "no_presentado");
  assert.equal(R.leerFilasExcel([["Nombre", "Edad"], ["Ana", 3]]), null);
});

test("analizarReservas calcula días, franjas, servicios y porcentajes", () => {
  const r = (fecha, hora, servicio, estado, reprogramada = "no") => ({ fecha, hora, matricula: "0000 BBB", servicio, estado, reprogramada });
  const reservas = [
    r("2026-11-02", "09:00", "Frenos", "entregada"), // lunes
    r("2026-11-02", "09:10", "Frenos", "cancelada", "sí"),
    r("2026-11-09", "09:30", "Revisión", "entregada"), // lunes
    r("2026-11-03", "16:00", "Batería", "rechazada"), // martes
    r("2026-11-06", "18:30", "Frenos", "no_presentado"), // viernes
    r("2026-11-06", "10:00", "Pre-ITV", "pendiente", "sí"),
    r("2026-11-05", "11:00", "Revisión", "confirmada"), // jueves
    r("2026-11-05", "11:00", "Revisión", "en_taller"),
    r("2026-11-05", "12:00", "Aceite y filtros", "lista"),
    r("2026-11-04", "13:30", "Diagnosis", "entregada"), // miércoles
  ];
  const a = R.analizarReservas(reservas);
  assert.equal(a.total, 10);
  assert.equal(a.desde, "2026-11-02");
  assert.equal(a.hasta, "2026-11-09");
  assert.deepEqual(a.diasSemana[0], { dia: "Lunes", reservas: 3 });
  assert.deepEqual(a.diasSemana[1], { dia: "Jueves", reservas: 3 });
  assert.equal(a.diasSemana.length, 5);
  assert.deepEqual(a.franjasMasLlenas[0], { franja: "09:00", reservas: 2 });
  assert.deepEqual(a.franjasMasLlenas[1], { franja: "11:00", reservas: 2 });
  assert.equal(a.franjasMasVacias[0].reservas, 0);
  assert.deepEqual(a.servicios[0], { servicio: "Frenos", reservas: 3, porcentaje: 30 });
  assert.deepEqual(a.servicios[1], { servicio: "Revisión", reservas: 3, porcentaje: 30 });
  assert.deepEqual(a.estados, {
    aceptadas: 60, // entregada x3, confirmada, en_taller, lista
    reprogramadas: 20,
    canceladas: 10,
    rechazadas: 10,
    no_presentadas: 10,
    pendientes: 10,
  });
  assert.equal(a.diasConReservas, 6);
  assert.equal(a.mediaPorDia, 1.7);
});

test("franjaDe agrupa en tramos de 30 minutos", () => {
  assert.equal(R.franjaDe("09:00"), "09:00");
  assert.equal(R.franjaDe("09:29"), "09:00");
  assert.equal(R.franjaDe("09:30"), "09:30");
  assert.equal(R.franjaDe("17:45"), "17:30");
});

test("a la IA solo van cifras agregadas: ni matrículas ni reservas sueltas", () => {
  const res = R.generarReservasDemo({ hoy: HOY });
  const datos = R.agregadosParaIA(R.analizarReservas(res));
  const texto = JSON.stringify(datos);
  for (const r of res) assert.ok(!texto.includes(r.matricula), "se cuela una matrícula");
  assert.deepEqual(Object.keys(datos).sort(), [
    "desde", "diasSemana", "franjasMasLlenas", "franjasMasVacias", "hasta", "mediaPorDia", "porcentajes", "servicios", "total",
  ]);
});

test("reservasPorDia rellena con 0 los días laborables sin reservas y no añade fines de semana", () => {
  const r = (fecha) => ({ fecha, hora: "09:00", matricula: "", servicio: "Otro", estado: "entregada", reprogramada: "no" });
  const dias = R.reservasPorDia([r("2026-11-06"), r("2026-11-06"), r("2026-11-10")]); // viernes y martes
  assert.deepEqual(dias, [
    { fecha: "2026-11-06", reservas: 2 },
    { fecha: "2026-11-09", reservas: 0 },
    { fecha: "2026-11-10", reservas: 1 },
  ]);
  assert.deepEqual(R.reservasPorDia([]), []);
});
