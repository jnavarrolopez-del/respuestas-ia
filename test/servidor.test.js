// Pruebas del servidor de recomendaciones con una IA simulada (no gasta saldo).
const test = require("node:test");
const assert = require("node:assert/strict");
const R = require("../src/reservas.js");
const S = require("../server/recomendaciones.js");

const agregados = R.agregadosParaIA(R.analizarReservas(R.generarReservasDemo({ hoy: new Date(2026, 10, 26) })));

function clienteFalso(respuesta, apuntar) {
  return {
    beta: {
      messages: {
        create: async (params) => {
          if (apuntar) apuntar(params);
          return respuesta;
        },
      },
    },
  };
}
const respuestaOK = (lista) => ({ stop_reason: "end_turn", content: [{ type: "text", text: JSON.stringify({ recomendaciones: lista }) }] });

test("acepta las cifras que genera el panel y las deja igual", () => {
  const limpio = S.validarAgregados(JSON.parse(JSON.stringify(agregados)));
  assert.equal(limpio.total, 150);
  assert.deepEqual(limpio.diasSemana, agregados.diasSemana);
  assert.deepEqual(limpio.franjasMasLlenas, agregados.franjasMasLlenas);
  assert.deepEqual(Object.keys(limpio.porcentajes), ["aceptadas", "reprogramadas", "canceladas", "rechazadas", "no_presentadas", "pendientes"]);
});

test("rechaza datos con otra forma y quita campos de más (como reservas sueltas)", () => {
  assert.throws(() => S.validarAgregados(null), S.ErrorDatos);
  assert.throws(() => S.validarAgregados({ ...agregados, total: -1 }), S.ErrorDatos);
  assert.throws(() => S.validarAgregados({ ...agregados, diasSemana: [{ dia: "Lunesito", reservas: 1 }] }), S.ErrorDatos);
  assert.throws(() => S.validarAgregados({ ...agregados, servicios: new Array(13).fill({ servicio: "x", reservas: 1, porcentaje: 1 }) }), S.ErrorDatos);
  const conExtra = S.validarAgregados({ ...agregados, reservas: [{ matricula: "1234 BCD" }], matriculas: ["1234 BCD"] });
  assert.ok(!JSON.stringify(conExtra).includes("1234 BCD"));
  const largo = S.validarAgregados({ ...agregados, servicios: [{ servicio: "a\n".repeat(100), reservas: 1, porcentaje: 1 }] });
  assert.ok(largo.servicios[0].servicio.length <= 40 && !largo.servicios[0].servicio.includes("\n"));
});

test("pide a Claude con la clave del servidor y devuelve 2 o 3 recomendaciones", async () => {
  let enviado;
  const recs = await S.pedirRecomendaciones(S.validarAgregados(agregados), {
    apiKey: "clave-de-prueba",
    client: clienteFalso(respuestaOK(["Uno.", " Dos. ", "Tres.", "Cuatro."]), (p) => (enviado = p)),
  });
  assert.deepEqual(recs, ["Uno.", "Dos.", "Tres."]);
  assert.equal(enviado.model, "claude-opus-5-5");
  assert.ok(enviado.system.includes("tuteando"));
  assert.ok(!JSON.stringify(enviado.messages).match(/\d{4} [A-Z]{3}/), "una matrícula en el mensaje");
});

test("da error si la IA rechaza, se corta o devuelve menos de 2", async () => {
  const opc = (r) => ({ apiKey: "x", client: clienteFalso(r) });
  await assert.rejects(S.pedirRecomendaciones(agregados, opc({ stop_reason: "refusal", content: [] })));
  await assert.rejects(S.pedirRecomendaciones(agregados, opc({ stop_reason: "max_tokens", content: [] })));
  await assert.rejects(S.pedirRecomendaciones(agregados, opc(respuestaOK(["Solo una."]))));
});

test("sin ANTHROPIC_API_KEY explica que falta la clave en el servidor", async () => {
  const antes = process.env.ANTHROPIC_API_KEY;
  delete process.env.ANTHROPIC_API_KEY;
  try {
    await assert.rejects(S.pedirRecomendaciones(agregados), (err) => {
      assert.deepEqual(S.respuestaDeError(err), { status: 500, mensaje: "El servidor no tiene configurada la clave de la IA." });
      return true;
    });
  } finally {
    if (antes !== undefined) process.env.ANTHROPIC_API_KEY = antes;
  }
});

test("el limitador deja 10 peticiones por minuto y por IP", () => {
  const permitir = S.crearLimitador(10);
  for (let i = 0; i < 10; i++) assert.ok(permitir("1.2.3.4"));
  assert.equal(permitir("1.2.3.4"), false);
  assert.ok(permitir("5.6.7.8"));
});
