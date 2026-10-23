const SHEET_ID = '1-8grei5ZHZAWGTtR0Ti6mc9L1DRZgaTvrqBNrTjaamw';
const CAMPOS = ['fecha', 'nombre', 'celular', 'correo', 'servicio', 'servicio_otro',
  'zona', 'zona_otra', 'pregunta', 'consentimiento', 'honeypot'];

const SERVICIOS = ['Electricidad', 'Pintura', 'Limpieza del hogar', 'Albañilería / maestro de obra',
  'Carpintería', 'Cerrajería', 'Fletes y mudanzas', 'Jardinería',
  'Reparación de electrodomésticos (línea blanca)', 'Otro'];
const ZONAS = ['Norte', 'Sur', 'Valle de los Chillos', 'Cumbayá', 'Otra'];

function validar(data) {
  const nombre = String(data.nombre ?? '').trim();
  if (nombre.length < 2 || nombre.length > 100) return 'Revisa tu nombre.';

  const celular = String(data.celular ?? '').trim();
  if (!/^\+?[0-9 ]{7,20}$/.test(celular)) return 'Revisa tu número celular.';

  const correo = String(data.correo ?? '').trim();
  if (correo && (correo.length > 100 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo))) {
    return 'Revisa tu correo.';
  }

  if (!SERVICIOS.includes(data.servicio)) return 'Selecciona un servicio.';
  if (!ZONAS.includes(data.zona)) return 'Selecciona una zona.';
  if (data.consentimiento !== 'Sí') return 'Falta tu consentimiento.';
  if (String(data.pregunta ?? '').length > 500) return 'La pregunta es demasiado larga.';
  return '';
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.honeypot) return respuesta({ ok: true });
    const error = validar(data);
    if (error) return respuesta({ ok: false, error });

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    const fila = CAMPOS.map(campo => {
      if (campo === 'fecha') return new Date();
      return String(data[campo] ?? '').slice(0, 500);
    });
    const filaNueva = sheet.getLastRow() + 1;
    sheet.appendRow(fila);
    sheet.getRange(filaNueva, CAMPOS.indexOf('celular') + 1)
      .setNumberFormat('@')
      .setValue(String(data.celular ?? ''));
    return respuesta({ ok: true });
  } catch (err) {
    return respuesta({ ok: false, error: String(err) });
  }
}

function formatearHoja() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  sheet.getRange(1, 1, 1, CAMPOS.length)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#7809D9')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.setColumnWidths(1, CAMPOS.length, 150);
  sheet.setColumnWidth(9, 300);
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
