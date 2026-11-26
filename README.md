# Análisis de reservas (Cuotivo) · complemento de Excel

Panel de Excel para analizar las reservas de un taller mecánico. Esta versión **no se conecta a Cuotivo**: trabaja con un taller de demostración.

## Qué hace

1. **Cargar taller de demostración**: crea una hoja «Reservas demo» con unas 150 reservas inventadas de los últimos 2 meses:
   - columnas: fecha, hora de entrega, matrícula, servicio, estado y reprogramada (sí/no);
   - de lunes a viernes, en jornada partida (9:00–14:00 y 16:00–19:00) y en franjas de 30 minutos.
2. **Analizar reservas**: calcula **con código** (sin IA):
   - los días de la semana con más reservas;
   - las 3 franjas más llenas y las 3 más vacías;
   - los servicios más pedidos;
   - el % de reservas aceptadas (confirmada, en taller, lista o entregada), reprogramadas, canceladas, rechazadas, no presentadas y pendientes.

   Después envía **solo esas cifras agregadas** al servidor, que pide a Claude 2 o 3 recomendaciones cortas para el dueño, en español y tuteando. Nunca se envían matrículas ni reservas sueltas.
3. **Crear gráfico**: crea la hoja «Gráfico de reservas» con la tabla de reservas por día y un gráfico de columnas.

Analiza la hoja abierta si tiene las columnas Fecha, Hora de entrega, Servicio y Estado. Si no, usa la última hoja de demostración.

## La clave de Claude va en el servidor

El navegador nunca ve la clave. El panel llama a `POST /api/recomendaciones` y el servidor:

- comprueba que el cuerpo tiene la forma exacta de las cifras agregadas y descarta cualquier otro campo;
- llama a Claude (`claude-opus-5-5`) con la clave guardada en `ANTHROPIC_API_KEY`;
- limita las peticiones a 10 por minuto y por IP.

Hay dos servidores con la misma lógica (`server/recomendaciones.js`):

- **Local**: `server.js` sirve el panel en `https://localhost:3000` y lee la clave de `.env`.
- **Vercel**: `api/recomendaciones.js`. La clave se guarda en las variables de entorno del proyecto.

GitHub Pages solo sirve archivos, no ejecuta servidores. Con la versión publicada allí funcionan la demostración, el análisis y el gráfico, pero **no las recomendaciones**: el panel lo avisa.

## Configurar y probar en local

1. Instala las dependencias (una vez):

   ```powershell
   npm install
   ```

2. Crea el archivo `.env` en esta carpeta, con tu clave de Anthropic. Está en `.gitignore`, así que no se sube a GitHub:

   ```
   ANTHROPIC_API_KEY=sk-ant-...
   ```

3. Si no los tienes, instala los certificados de desarrollo de Office (una vez):

   ```powershell
   npx office-addin-dev-certs install
   ```

4. Arranca el servidor:

   ```powershell
   npm start
   ```

5. En Excel web: **Complementos → Más complementos → Administrar mis complementos ▾ → Cargar mi complemento** → `manifest-local.xml`.
6. En **Inicio**, pulsa **Análisis de reservas** y sigue los tres botones en orden.

## Publicarlo con recomendaciones (Vercel)

1. Importa este repositorio en Vercel. No necesita comando de compilación.
2. En **Settings → Environment Variables**, añade `ANTHROPIC_API_KEY`.
3. En `manifest.xml`, cambia `https://jnavarrolopez-del.github.io/respuestas-ia/src` por `https://TU-PROYECTO.vercel.app/src`. Mantén el mismo `<Id>`.
4. Vuelve a cargar `manifest.xml` en Excel.

## Pruebas

```powershell
npm test
```

Prueban la generación de reservas, el análisis, la lectura y escritura de las hojas, y el servidor (con una IA simulada, sin gastar saldo).

## Archivos

- `manifest.xml` (publicado) y `manifest-local.xml` (local): mismo ID que la versión anterior del complemento.
- `src/taskpane.html` y `src/taskpane.js`: el panel y las operaciones de Excel (hojas, tabla y gráfico).
- `src/reservas.js`: generador de demostración y análisis, sin dependencias.
- `server/recomendaciones.js`: validación de las cifras y llamada a Claude.
- `server.js`: servidor local. `api/recomendaciones.js`: función para Vercel.
