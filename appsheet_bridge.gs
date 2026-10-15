const AUDIT_SHEET = '1UTRfVnqYaG7Qw3UBF4YwAcrgir36ul8zOgOtED46LG0';

// Run once as owner. The token stays in Script Properties, never in source.
function configurarConexion() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('AUDIT_TOKEN')) {
    props.setProperty('AUDIT_TOKEN', Utilities.getUuid() + Utilities.getUuid());
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Solicitud vacía');
    const body = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('AUDIT_TOKEN');
    if (!token || body.token !== token) return respuesta({ok:false,error:'No autorizado'});
    if (!lock.tryLock(25000)) return respuesta({ok:false,error:'Ocupado; reintentar'});
    const ss = SpreadsheetApp.openById(AUDIT_SHEET);
    if (body.action === 'read') {
      const data = {};
      ['Pedidos','Articulos','Validaciones','Auditorias','Diferencias'].forEach(name => {
        data[name] = registros(ss.getSheetByName(name));
      });
      return respuesta({ok:true,data:data});
    }
    if (body.action === 'publish') {
      ['Pedidos','Articulos'].forEach(name => {
        if (Array.isArray(body[name]) && body[name].length) guardar(ss.getSheetByName(name),body[name]);
      });
      return respuesta({ok:true});
    }
    if (body.action === 'audit') {
      const audit = body.audit;
      if (!audit || !audit.ID_Auditoria || !audit.ID_Pedido) throw new Error('Auditoría inválida');
      if (!registros(ss.getSheetByName('Pedidos')).some(r => String(r.ID_Pedido) === String(audit.ID_Pedido))) {
        throw new Error('Pedido no publicado');
      }
      if (body.pdf_base64) {
        const bytes = Utilities.base64Decode(body.pdf_base64);
        if (bytes.length > 10000000) throw new Error('PDF demasiado grande');
        const parents = DriveApp.getFileById(AUDIT_SHEET).getParents();
        const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
        const folders = parent.getFoldersByName('Auditoria_PDF');
        const folder = folders.hasNext() ? folders.next() : parent.createFolder('Auditoria_PDF');
        const fileName = 'auditoria_' + String(audit.ID_Auditoria).replace(/[^a-zA-Z0-9_-]/g,'') + '.pdf';
        if (!folder.getFilesByName(fileName).hasNext()) {
          folder.createFile(Utilities.newBlob(bytes,'application/pdf',fileName));
        }
        audit.PDF_Archivo = 'Auditoria_PDF/' + fileName;
      }
      guardar(ss.getSheetByName('Auditorias'),[audit]);
      if (body.Diferencias) guardar(ss.getSheetByName('Diferencias'),body.Diferencias);
      return respuesta({ok:true,pdf:audit.PDF_Archivo || ''});
    }
    throw new Error('Operación no permitida');
  } catch (err) {
    return respuesta({ok:false,error:String(err.message || err)});
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function respuesta(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
function registros(sheet) {
  if (!sheet) throw new Error('Tabla ausente');
  const values = sheet.getDataRange().getValues();
  const headers = values.length ? values.shift() : [];
  return values.filter(row => row[0] !== '').map(row => {
    const record = {};
    headers.forEach((name,i) => { if (name) record[name] = row[i] == null ? '' : row[i]; });
    return record;
  });
}
function guardar(sheet, records) {
  if (!sheet || !Array.isArray(records) || records.length > 20000) throw new Error('Carga inválida');
  if (!records.length) return;
  const width = sheet.getLastColumn();
  if (!width) throw new Error('La tabla ' + sheet.getName() + ' no tiene encabezados');
  const headers = sheet.getRange(1,1,1,width).getValues()[0];
  if (!headers[0]) throw new Error('Falta la columna identificadora en ' + sheet.getName());
  const count = Math.max(0,sheet.getLastRow() - 1);
  const existing = count ? sheet.getRange(2,1,count,width).getValues() : [];
  const rows = existing.map(row => Array.from({length:width},(_,i) => row[i] == null ? '' : row[i]));
  const indices = new Map(rows.map((row,i) => [String(row[0]),i]));
  records.forEach(record => {
    if (!record || typeof record !== 'object' || Array.isArray(record)) {
      throw new Error('Registro inválido en ' + sheet.getName());
    }
    const key = String(record[headers[0]] == null ? '' : record[headers[0]]).trim();
    if (!key) throw new Error('Identificador vacío en ' + sheet.getName());
    let at = indices.has(key) ? indices.get(key) : rows.length;
    const current = rows[at] || Array(width).fill('');
    headers.forEach((name,i) => {
      if (name && Object.prototype.hasOwnProperty.call(record,name)) {
        let value = record[name] == null ? '' : record[name];
        if (typeof value === 'string' && /^[=+@]/.test(value)) value = "'" + value;
        current[i] = value;
      }
    });
    rows[at] = Array.from({length:width},(_,i) => current[i] == null ? '' : current[i]);
    indices.set(key,at);
  });
  if (!rows.length) return;
  if (sheet.getMaxRows() < rows.length + 1) {
    sheet.insertRowsAfter(sheet.getMaxRows(),rows.length + 1 - sheet.getMaxRows());
  }
  const batchSize = 500;
  for (let offset = 0; offset < rows.length; offset += batchSize) {
    const batch = rows.slice(offset,offset + batchSize);
    try {
      sheet.getRange(offset + 2,1,batch.length,width).setValues(batch);
    } catch (err) {
      throw new Error('Error al guardar ' + sheet.getName() + ' (filas ' + (offset + 2) + '–' + (offset + 1 + batch.length) + ', columnas: ' + width + '): ' + String(err.message || err));
    }
  }
}
