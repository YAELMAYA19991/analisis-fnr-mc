const AUDIT_SHEET = '1UTRfVnqYaG7Qw3UBF4YwAcrgir36ul8zOgOtED46LG0';

// Run once as owner. The token stays in Script Properties, never in source.
function configurarConexion() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('AUDIT_TOKEN')) {
    props.setProperty('AUDIT_TOKEN', Utilities.getUuid() + Utilities.getUuid());
  }
}

// Run manually in the Apps Script editor if Streamlit needs the existing token.
// The execution log is visible only to project editors; never publish this value.
function mostrarTokenConexion() {
  const token = PropertiesService.getScriptProperties().getProperty('AUDIT_TOKEN');
  if (!token) throw new Error('Primero ejecuta configurarConexion().');
  Logger.log('Copia este valor a APPSHEET_BRIDGE_TOKEN en Streamlit Secrets: ' + token);
}

// Read-only health check; it exposes no sheet data or token.
function doGet(e) {
  return respuesta({ok:true,service:'appsheet-audit-bridge',message:'Endpoint activo; POST protegido por token.'});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Solicitud vacía');
    const body = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('AUDIT_TOKEN');
    if (!token || body.token !== token) return respuesta({ok:false,error:'No autorizado'});
    if (body.action === 'health') {
      const ss = SpreadsheetApp.openById(AUDIT_SHEET);
      const required = ['Pedidos','Articulos','Validaciones','Auditorias','Diferencias'];
      const missing = required.filter(name => !ss.getSheetByName(name));
      if (missing.length) throw new Error('Faltan hojas requeridas: ' + missing.join(', '));
      const counts = {};
      required.forEach(name => {
        counts[name] = Math.max(0, ss.getSheetByName(name).getLastRow() - 1);
      });
      return respuesta({ok:true,message:'Puente y hoja conectados',tables:counts,capabilities:{replaceBatch:true,serverFingerprint:true}});
    }
    if (!lock.tryLock(25000)) return respuesta({ok:false,error:'Ocupado; reintentar'});
    const ss = SpreadsheetApp.openById(AUDIT_SHEET);
    if (body.action === 'read') {
      const data = {};
      ['Pedidos','Articulos','Validaciones','Auditorias','Diferencias'].forEach(name => {
        data[name] = registros(ss.getSheetByName(name));
      });
      return respuesta({ok:true,data:data,fingerprint:huellaLote(data)});
    }
    if (body.action === 'publish') {
      if (body.replaceBatch === true) {
        if (body.archiveConfirmed !== true) throw new Error('El reemplazo requiere confirmar el archivado previo.');
        const current = {};
        ['Pedidos','Articulos','Validaciones','Auditorias','Diferencias'].forEach(name => {
          current[name] = registros(ss.getSheetByName(name));
        });
        if (!body.expectedFingerprint || body.expectedFingerprint !== huellaLote(current)) {
          throw new Error('El lote cambió desde la lectura; no se borró ningún dato. Vuelve a cargar y reintenta.');
        }
        validarNuevoLote(ss,body);
        ['Pedidos','Articulos','Validaciones','Auditorias','Diferencias'].forEach(name => limpiarDatos(ss.getSheetByName(name)));
        formatoCodigoTexto(ss.getSheetByName('Articulos'));
        guardar(ss.getSheetByName('Pedidos'),body.Pedidos);
        guardar(ss.getSheetByName('Articulos'),body.Articulos);
        return respuesta({ok:true,replaced:true,message:'Lote anterior archivado y lote nuevo cargado.'});
      }
      ['Pedidos','Articulos'].forEach(name => {
        if (Array.isArray(body[name]) && body[name].length) guardar(ss.getSheetByName(name),body[name]);
      });
      return respuesta({ok:true,replaced:false});
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

function huellaLote(tables) {
  const names = ['Pedidos','Articulos','Validaciones','Auditorias','Diferencias'];
  const payload = {};
  names.forEach(name => {
    const rows = (tables[name] || []).map(row => {
      const normalized = {};
      Object.keys(row || {}).sort().forEach(key => {
        const value = row[key];
        normalized[key] = value == null ? '' : value;
      });
      return JSON.stringify(normalized);
    });
    rows.sort();
    payload[name] = rows;
  });
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(payload),
    Utilities.Charset.UTF_8
  );
  return digest.map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
}

function validarNuevoLote(ss, body) {
  const orders = body.Pedidos;
  const articles = body.Articulos;
  if (!Array.isArray(orders) || !orders.length || orders.length > 20000) throw new Error('El lote nuevo no contiene pedidos válidos.');
  if (!Array.isArray(articles) || !articles.length || articles.length > 20000) throw new Error('El lote nuevo no contiene artículos válidos.');
  const orderSheet = ss.getSheetByName('Pedidos');
  const articleSheet = ss.getSheetByName('Articulos');
  const orderHeaders = orderSheet.getRange(1,1,1,orderSheet.getLastColumn()).getValues()[0];
  const articleHeaders = articleSheet.getRange(1,1,1,articleSheet.getLastColumn()).getValues()[0];
  ['ID_Pedido','Slot'].forEach(name => { if (orderHeaders.indexOf(name) < 0) throw new Error('Falta la columna ' + name + ' en Pedidos.'); });
  ['ID_Articulo','ID_Pedido','Articulo','Picker','Cantidad_Pedida','Cantidad_Pickeada','Codigo_Barras','Zona'].forEach(name => {
    if (articleHeaders.indexOf(name) < 0) throw new Error('Falta la columna ' + name + ' en Articulos.');
  });
  const orderIds = new Set();
  orders.forEach(row => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) throw new Error('Registro inválido en Pedidos.');
    const id = String(row.ID_Pedido || '').trim();
    if (!id || !String(row.Slot || '').trim() || orderIds.has(id)) throw new Error('Pedido incompleto o duplicado.');
    orderIds.add(id);
  });
  const articleIds = new Set();
  articles.forEach(row => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) throw new Error('Registro inválido en Articulos.');
    const id = String(row.ID_Articulo || '').trim();
    const order = String(row.ID_Pedido || '').trim();
    const code = String(row.Codigo_Barras || '').trim();
    const zone = String(row.Zona || '').trim();
    const picker = String(row.Picker || '').trim();
    if (!id || articleIds.has(id) || !orderIds.has(order) || !code || !zone ||
        !String(row.Articulo || '').trim() || !String(row.Cantidad_Pedida || '').trim() ||
        !String(row.Cantidad_Pickeada || '').trim() || !picker || picker.indexOf('@') < 1) {
      throw new Error('Artículo incompleto, duplicado o ligado a un pedido inexistente.');
    }
    articleIds.add(id);
  });
}

function limpiarDatos(sheet) {
  if (!sheet) throw new Error('Tabla ausente durante el reemplazo.');
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow > 1 && lastColumn > 0) sheet.getRange(2,1,lastRow - 1,lastColumn).clearContent();
}

function formatoCodigoTexto(sheet) {
  if (!sheet) throw new Error('Tabla Articulos ausente.');
  const width = sheet.getLastColumn();
  if (!width) throw new Error('Tabla Articulos sin encabezados.');
  const headers = sheet.getRange(1,1,1,width).getValues()[0];
  const codeColumn = headers.indexOf('Codigo_Barras') + 1;
  if (codeColumn < 1) throw new Error('Falta Codigo_Barras en Articulos.');
  sheet.getRange(2,codeColumn,Math.max(1,sheet.getMaxRows() - 1),1).setNumberFormat('@');
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
