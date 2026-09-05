/**
 * BRASAS — backend de pedidos (Google Apps Script)
 *
 * QUÉ HACE:
 * - Recibe cada pedido que un cliente arma en la web (doPost) y lo guarda como
 *   una fila nueva en esta misma planilla de Google.
 * - Le permite a la web leer los pedidos guardados (doGet) para mostrar el
 *   historial en el panel de administración ("📋 Ver historial de pedidos").
 * - Opcional: te manda un email cada vez que entra un pedido nuevo.
 *
 * CÓMO INSTALARLO (una sola vez):
 * 1. Andá a https://sheets.google.com y creá una planilla nueva (vacía).
 * 2. Arriba, andá a Extensiones → Apps Script.
 * 3. Borrá todo el código de ejemplo que aparece, y pegá TODO este archivo.
 * 4. Reemplazá el email de abajo (EMAIL_AVISO) por el tuyo, o dejalo vacío
 *    entre comillas ("") si no querés recibir mails.
 * 5. Arriba a la derecha, botón "Implementar" → "Nueva implementación".
 * 6. Tipo: "Aplicación web". Ejecutar como: "Yo". Quién tiene acceso: "Cualquier usuario".
 * 7. Autorizá los permisos que te pida Google (es tu propia cuenta, es seguro).
 * 8. Copiá la URL que te da ("URL de la aplicación web", termina en /exec).
 * 9. Pegala en la web, en el panel admin → "⚙️ Configuración" → "URL de la
 *    planilla de pedidos".
 *
 * Listo. Cada pedido nuevo va a aparecer como fila en esta planilla,
 * y vas a poder verlos también desde "📋 Ver historial de pedidos" en la web.
 */

const EMAIL_AVISO = ""; // poné tu email entre comillas si querés que te avise por mail, ej: "vos@gmail.com"
const NOMBRE_HOJA = "Pedidos";

function doPost(e) {
  const hoja = obtenerHoja();
  const datos = JSON.parse(e.postData.contents);
  hoja.appendRow([
    datos.fecha || new Date().toISOString(),
    datos.telefono || "",
    datos.resumen || "",
    datos.total || 0,
    datos.notaGeneral || "",
  ]);
  if (EMAIL_AVISO) {
    try {
      MailApp.sendEmail(
        EMAIL_AVISO,
        `Nuevo pedido BRASAS — ${datos.telefono}`,
        `Teléfono: ${datos.telefono}\nPedido:\n${datos.resumen}\n\nTotal: $${datos.total}\n\nNota: ${datos.notaGeneral || "-"}`
      );
    } catch (err) { /* si falla el mail, no importa, el pedido ya quedo guardado */ }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const hoja = obtenerHoja();
  const filas = hoja.getDataRange().getValues();
  filas.shift(); // saca la fila de encabezados
  const pedidos = filas.map(f => ({
    fecha: f[0], telefono: f[1], resumen: f[2], total: f[3], notaGeneral: f[4],
  }));
  return ContentService.createTextOutput(JSON.stringify(pedidos)).setMimeType(ContentService.MimeType.JSON);
}

function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
    hoja.appendRow(["Fecha", "Teléfono", "Resumen del pedido", "Total", "Nota"]);
  }
  return hoja;
}
