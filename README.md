# Respuestas IA para Excel

Copia una celda con una pregunta y pégala en una celda vacía: la respuesta aparece sola. Sin paneles ni avisos.

## Cómo funciona

1. Al **seleccionar** celdas con preguntas (para copiarlas), el complemento guarda esos textos solo en memoria.
2. Al **seleccionar el destino**, toma una foto local de esas celdas: si están vacías, su formato numérico y si tienen lista desplegable. Esta foto no sale del equipo.
3. Al **pegar**, cada celda que estaba vacía y ahora contiene una de esas preguntas se envía a la IA (solo el texto de la pregunta y el tipo de respuesta esperado). La respuesta sustituye a la pregunta y se recupera el formato numérico y la lista desplegable que tenía la celda.

| Celda de destino | Respuesta escrita |
|---|---|
| Formato numérico (`0`, `0,00`, `#.##0 €`), o encabezado de columna tipo «Cantidad», «Precio», «Año» | Número |
| Formato de porcentaje | Porcentaje (0,25 → 25 %) |
| Formato de fecha | Fecha |
| Lista desplegable | Una de las opciones de la lista (o nada, si la respuesta no coincide con ninguna) |
| Formato Texto (`@`) | Texto breve |
| General | Número si la respuesta es un número; la letra de la opción si es tipo test; si no, texto breve |

**Varias preguntas:** copia, por ejemplo, `A2:A10` y pega en `B2`. Cada celda recibe la respuesta de la pregunta que cayó en ella, en el mismo orden.

**Estados:**
- No es una pregunta, o la celda ya tenía contenido propio: no se hace nada.
- Mientras responde la IA: la celda muestra la pregunta pegada hasta que llega la respuesta.
- Error (sin conexión, la IA no lo sabe, la respuesta no encaja con el tipo de la celda): no se escribe nada y se añade una nota en la celda. Pega de nuevo en esa celda para reintentar; si sale bien, la nota desaparece.

## Requisitos

- Excel para Windows o Mac (Microsoft 365), o Excel en la web. Hace falta compatibilidad con el runtime compartido (*SharedRuntime 1.1*) y *ExcelApi 1.9*. Las notas de error usan *ExcelApi 1.18*; en versiones anteriores se usa un comentario.
- Node.js (para el servidor local).
- Una clave de API de Anthropic. Las respuestas se cobran en tu cuenta de Anthropic. Se usa el modelo `claude-opus-5-5`, que puedes cambiar en `src/addin.js` (`MODEL`).

## Instalación (una vez)

1. Certificado HTTPS local para Office. Esta orden instala un certificado de desarrollo de confianza en Windows; revísala antes de ejecutarla:

   ```powershell
   npx office-addin-dev-certs install
   ```

2. Arranca el servidor, que debe estar en marcha mientras uses Excel:

   ```powershell
   node server.js
   ```

3. Carga el complemento en Excel:
   - **Excel en la web:** Inicio > Complementos > Más complementos > Mis complementos > *Cargar mi complemento* y elige `manifest.xml`.
   - **Excel para Windows:**

     ```powershell
     npx office-addin-debugging start manifest.xml desktop
     ```

     Otra opción es una carpeta compartida como catálogo de confianza (Archivo > Opciones > Centro de confianza > Catálogos de complementos de confianza).

4. En la pestaña **Inicio**, pulsa **Configurar clave**, pega tu clave y pulsa **Guardar clave**. El panel se cierra solo. A partir de ese momento el complemento se carga en segundo plano cada vez que abras un libro.

**Versión publicada:** `manifest.xml` apunta a GitHub Pages (`https://jnavarrolopez-del.github.io/respuestas-ia/src/`), así que no hace falta tener `node server.js` encendido. Para probar cambios en local usa `manifest-local.xml`, que apunta a `https://localhost:3000`.

## Privacidad

- La clave se guarda en el almacenamiento local del complemento en este equipo, no en el libro.
- A la IA solo se envía el texto de la pregunta y el tipo esperado (`number`, `percent`, `date`, `option`, `short_text` o `auto`). Ni encabezados, ni opciones de listas, ni otras celdas.
- La última pregunta copiada y las fotos de las celdas de destino solo existen en memoria y se pierden al cerrar Excel.

## Limitaciones

- Office no deja leer el portapapeles. Por eso el complemento identifica lo copiado como la **última selección de la hoja que contenía preguntas**, y lo reconoce al pegarlo. No funciona con preguntas copiadas desde otro programa u otro libro.
- Las preguntas se detectan con una regla local: llevan `?`/`¿`, empiezan por una palabra interrogativa o una orden («Calcula…», «Indica…»), o tienen opciones `a) … b) …`.
- Un pegado normal sobrescribe también colores y bordes del destino. El complemento solo recupera el formato numérico y la lista desplegable. Para conservar todo el formato, usa *Pegar valores*.
- Si el destino tenía contenido, Excel ya lo ha sobrescrito al pegar. El complemento no lo toca más, y Ctrl+Z lo recupera.
- Como máximo 400 celdas por pegado.

## Archivos

- `manifest.xml`: manifiesto del complemento, con runtime compartido de vida larga.
- `src/taskpane.html` y `src/taskpane.js`: pantalla de configuración de la clave y arranque en segundo plano.
- `src/addin.js`: eventos de Excel, llamada a la IA y escritura de las respuestas.
- `src/logic.js`: reglas sin dependencias (detección, tipos, conversión de valores), con pruebas en `test/`.
- `server.js`: servidor HTTPS local sin dependencias.

Pruebas:

```powershell
node --test test/logic.test.js
```
