/**
 * APPS SCRIPT — Unidades Regionales (URE)
 * SOLO LECTURA: no hay doPost ni clave, porque personal y vehículos
 * se editan directamente en la planilla de Google Sheets.
 *
 * Qué hace:
 *  - GET → devuelve el contenido de las 3 pestañas ("Personal", "Vehiculos",
 *    "Regionales") en un solo JSON: { ok:true, personal:[...], vehiculos:[...], regionales:[...] }
 *
 * CÓMO USARLO:
 *  1. Abrí la planilla de Google Sheets que tiene las 3 pestañas
 *     "Personal", "Vehiculos" y "Regionales" (los nombres deben ser
 *     EXACTAMENTE esos, sin tildes, respetando mayúsculas/minúsculas).
 *  2. Extensiones → Apps Script (o abrí el proyecto que ya tenías creado).
 *  3. Borrá el contenido de Code.gs y pegá TODO este archivo.
 *  4. Guardá (ícono de disquete).
 *  5. Implementar → Nueva implementación → tipo "Aplicación web".
 *       - Ejecutar como: Yo (tu cuenta)
 *       - Quién tiene acceso: Cualquier usuario
 *  6. Autorizá los permisos cuando te lo pida Google.
 *  7. Copiá la URL que te da ("URL de la aplicación web") — esa es la
 *     que hay que pegar en unidades_regionales.html, en la constante
 *     URL_APPS_SCRIPT_URE.
 *
 * Si en el futuro editás este código, hay que volver a "Implementar" →
 * "Gestionar implementaciones" → editar → Nueva versión, para que los
 * cambios se apliquen a la URL ya publicada.
 */

const HOJAS = {
  personal: "Personal",
  vehiculos: "Vehiculos",
  regionales: "Regionales"
};

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultado = {
      ok: true,
      personal: leerHojaComoObjetos(ss, HOJAS.personal),
      vehiculos: leerHojaComoObjetos(ss, HOJAS.vehiculos),
      regionales: leerHojaComoObjetos(ss, HOJAS.regionales)
    };
    return ContentService
      .createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function leerHojaComoObjetos(ss, nombreHoja) {
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) {
    throw new Error('No se encontró la pestaña "' + nombreHoja + '".');
  }
  const valores = hoja.getDataRange().getValues();
  if (valores.length < 2) return [];

  const encabezados = valores[0];
  const filas = valores.slice(1);

  return filas
    .filter(fila => fila.some(celda => celda !== "" && celda !== null))
    .map(fila => {
      const obj = {};
      encabezados.forEach((col, i) => {
        let v = fila[i];
        if (Object.prototype.toString.call(v) === "[object Date]") {
          v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        obj[col] = v === "" || v === null || v === undefined ? null : v;
      });
      return obj;
    });
}
