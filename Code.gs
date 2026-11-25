/************************************************************
 * AEST | TIENDA VIRTUAL
 * Proyecto vinculado a Google Sheets.
 *
 * CONFIGURACIÓN INICIAL:
 * 1) Cambia PRODUCT_SHEET_NAME por el nombre real de tu pestaña
 *    de productos. Si lo dejas vacío, se usa la primera pestaña
 *    que no sea CONFIGURACION.
 * 2) Conserva el ID de tu carpeta principal de Drive.
 ************************************************************/

const AEST = {
  STORE_NAME: 'AEST',
  SPREADSHEET_ID: '', // Vacío = usa el Spreadsheet vinculado.
  PRODUCT_SHEET_NAME: '', // Ejemplo: 'PRODUCTOS'. Vacío = detección automática.
  CONFIG_SHEET_NAME: 'CONFIGURACION',
  HEADER_ROW: 1,
  DATA_START_ROW: 2,
  DRIVE_ROOT_ID: '1U7NJyog2OfQINuCr5_0TWsKJufTjKhwP',
  WHATSAPP: '51974407321',
  CACHE_SECONDS: 600,
  MAX_IMAGE_BYTES: 6 * 1024 * 1024,
  FOLDERS: {
    LOGO: 'AEST_LOGO',
    COVER: 'AEST_PORTADA',
    PRODUCTS: 'AEST_PRODUCTOS'
  },
  HEADERS: {
    PRODUCT: 'PRODUCTO',
    SIZE: 'TALLA',
    PRICE: 'PRECIO',
    IMAGE: 'PRODUCTO URL',
    STOCK: 'STOCK',
    GALLERY: 'PRODUCTO_GALERIA'
  }
};

/** Menú administrativo. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AEST | Administración')
    .addItem('Abrir panel de imágenes', 'abrirSidebarAEST')
    .addSeparator()
    .addItem('Preparar / revisar hojas y carpetas', 'configurarAEST')
    .addItem('Confirmar venta de fila activa', 'confirmarVentaFilaActiva')
    .addItem('Actualizar catálogo (limpiar caché)', 'limpiarCacheAEST')
    .addToUi();
}

/** Ejecutar una vez desde Apps Script y autorizar los permisos. */
function configurarAEST() {
  assertAdmin_();
  const ss = getSS_();
  const productSheet = getProductSheet_();
  ensureConfigSheet_();
  ensureProductColumns_(productSheet);
  ensureFolders_();
  limpiarCacheAEST();
  SpreadsheetApp.getUi().alert(
    'AEST está preparado. Revisa que la pestaña de productos tenga los encabezados PRODUCTO, TALLA, PRECIO, PRODUCTO URL y STOCK.'
  );
}

/** Publica el catálogo antiguo y expone una API privada para el proxy Netlify. */
function doGet(e) {
  const api = e && e.parameter ? String(e.parameter.api || '') : '';
  if (api === 'catalog') {
    return jsonOutputAEST_(obtenerCatalogoAEST());
  }
  if (api === 'validate') {
    try {
      const rawItems = e && e.parameter ? String(e.parameter.items || '[]') : '[]';
      const items = JSON.parse(rawItems);
      return jsonOutputAEST_(validarCarritoAEST(items));
    } catch (error) {
      return jsonOutputAEST_({ ok: false, message: 'No se pudieron validar los productos: ' + error.message });
    }
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('AEST | Streetwear')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Respuesta JSON consumida desde las funciones proxy de Netlify. */
function jsonOutputAEST_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Abre el panel privado dentro de Google Sheets. */
function abrirSidebarAEST() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AEST | Panel de imágenes')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Contexto de la fila seleccionada para el Sidebar. */
function obtenerContextoAdmin() {
  assertAdmin_();
  const ss = getSS_();
  const sheet = getProductSheet_();
  const active = ss.getActiveRange();
  let row = active ? active.getRow() : AEST.DATA_START_ROW;
  if (row < AEST.DATA_START_ROW) row = AEST.DATA_START_ROW;
  const map = getHeaderMap_(sheet);
  let product = '';
  if (map[AEST.HEADERS.PRODUCT]) {
    product = String(sheet.getRange(row, map[AEST.HEADERS.PRODUCT]).getDisplayValue() || '').trim();
  }
  const config = readConfig_();
  return {
    row: row,
    product: product,
    sheetName: sheet.getName(),
    logoUrl: config.LOGO_URL || '',
    coverUrl: config.PORTADA_URL || ''
  };
}

/** Devuelve solo los campos públicos necesarios para la tienda. */
function obtenerCatalogoAEST() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('AEST_PUBLIC_CATALOG_V1');
  if (cached) return JSON.parse(cached);

  const sheet = getProductSheet_();
  const map = getHeaderMap_(sheet);
  requireHeaders_(map, [AEST.HEADERS.PRODUCT, AEST.HEADERS.SIZE, AEST.HEADERS.PRICE, AEST.HEADERS.IMAGE, AEST.HEADERS.STOCK]);

  const lastRow = sheet.getLastRow();
  const config = readConfig_();
  const rows = lastRow >= AEST.DATA_START_ROW
    ? sheet.getRange(AEST.DATA_START_ROW, 1, lastRow - AEST.HEADER_ROW, sheet.getLastColumn()).getValues()
    : [];

  const products = [];
  rows.forEach(function(row, index) {
    const name = String(row[map[AEST.HEADERS.PRODUCT] - 1] || '').trim();
    if (!name) return;
    const size = String(row[map[AEST.HEADERS.SIZE] - 1] || '').trim();
    const price = parsePrice_(row[map[AEST.HEADERS.PRICE] - 1]);
    const stock = Math.max(0, Number(row[map[AEST.HEADERS.STOCK] - 1]) || 0);
    const rawImage = String(row[map[AEST.HEADERS.IMAGE] - 1] || '').trim();
    const gallery = map[AEST.HEADERS.GALLERY]
      ? String(row[map[AEST.HEADERS.GALLERY] - 1] || '').trim()
      : '';
    products.push({
      id: String(AEST.DATA_START_ROW + index),
      producto: name,
      talla: size,
      precio: price,
      stock: stock,
      imagenUrl: driveThumbnail_(rawImage, 900),
      galeria: gallery ? gallery.split('|').map(function(u) { return driveThumbnail_(u, 900); }).filter(Boolean) : []
    });
  });

  const result = {
    tienda: {
      nombre: config.NOMBRE_TIENDA || AEST.STORE_NAME,
      logoUrl: driveThumbnail_(config.LOGO_URL || '', 400),
      portadaUrl: driveThumbnail_(config.PORTADA_URL || '', 1800),
      whatsapp: normalizePhone_(config.WHATSAPP || AEST.WHATSAPP),
      presentacion: config.PRESENTACION || 'Poleras oversize · Streetwear'
    },
    productos: products
  };
  cache.put('AEST_PUBLIC_CATALOG_V1', JSON.stringify(result), AEST.CACHE_SECONDS);
  return result;
}

/**
 * Valida disponibilidad al intentar enviar el pedido.
 * No descuenta stock: el pedido se coordina por WhatsApp.
 */
function validarCarritoAEST(items) {
  if (!Array.isArray(items) || !items.length) {
    return { ok: false, message: 'Tu carrito está vacío.' };
  }
  const sheet = getProductSheet_();
  const map = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  const rows = lastRow >= AEST.DATA_START_ROW
    ? sheet.getRange(AEST.DATA_START_ROW, 1, lastRow - AEST.HEADER_ROW, sheet.getLastColumn()).getValues()
    : [];
  const errors = [];

  items.forEach(function(item) {
    const rowNumber = Number(item.row);
    const qty = Number(item.cantidad);
    if (!Number.isInteger(rowNumber) || rowNumber < AEST.DATA_START_ROW || !Number.isInteger(qty) || qty < 1) {
      errors.push('Hay un artículo con datos inválidos.');
      return;
    }
    const row = rows[rowNumber - AEST.DATA_START_ROW];
    if (!row) {
      errors.push('Uno de los productos ya no existe.');
      return;
    }
    const name = String(row[map[AEST.HEADERS.PRODUCT] - 1] || '').trim();
    const size = String(row[map[AEST.HEADERS.SIZE] - 1] || '').trim();
    const stock = Math.max(0, Number(row[map[AEST.HEADERS.STOCK] - 1]) || 0);
    if (name !== String(item.producto) || size !== String(item.talla)) {
      errors.push('La información de ' + String(item.producto || 'un producto') + ' cambió. Actualiza el catálogo.');
    } else if (qty > stock) {
      errors.push(name + ' talla ' + size + ': quedan ' + stock + ' unidades disponibles.');
    }
  });
  return errors.length ? { ok: false, message: errors.join('\n') } : { ok: true };
}

/** Sube logo, portada o fotos del modelo activo. */
function subirImagenesAEST(payload) {
  assertAdmin_();
  if (!payload || !payload.tipo || !Array.isArray(payload.files) || !payload.files.length) {
    throw new Error('Selecciona al menos una imagen.');
  }
  const folders = ensureFolders_();
  let folder;
  if (payload.tipo === 'logo') folder = folders.logo;
  else if (payload.tipo === 'portada') folder = folders.cover;
  else if (payload.tipo === 'producto') folder = folders.products;
  else throw new Error('Tipo de imagen no válido.');

  const uploaded = [];
  payload.files.forEach(function(fileData, index) {
    if (!fileData || !/^image\//i.test(String(fileData.mimeType || ''))) {
      throw new Error('Solo se permiten archivos de imagen.');
    }
    const bytes = Utilities.base64Decode(String(fileData.base64 || ''));
    if (bytes.length > AEST.MAX_IMAGE_BYTES) {
      throw new Error('Cada imagen debe pesar como máximo 6 MB: ' + String(fileData.name || 'imagen'));
    }
    const safeName = sanitizeFileName_(fileData.name || ('imagen_' + (index + 1)));
    const blob = Utilities.newBlob(bytes, fileData.mimeType, safeName);
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingError) {
      // Algunas cuentas de Google Workspace bloquean el acceso público por política.
      console.warn('No se pudo habilitar acceso público: ' + sharingError.message);
    }
    uploaded.push({
      id: file.getId(),
      url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200',
      name: file.getName()
    });
  });

  if (payload.tipo === 'logo') {
    writeConfigValue_('LOGO_URL', uploaded[0].url);
  } else if (payload.tipo === 'portada') {
    writeConfigValue_('PORTADA_URL', uploaded[0].url);
  } else {
    const productName = String(payload.producto || '').trim();
    if (!productName) throw new Error('No se detectó el nombre del producto. Selecciona una fila con PRODUCTO.');
    const sheet = getProductSheet_();
    const map = getHeaderMap_(sheet);
    const lastRow = sheet.getLastRow();
    const names = sheet.getRange(AEST.DATA_START_ROW, map[AEST.HEADERS.PRODUCT], Math.max(0, lastRow - AEST.HEADER_ROW), 1).getValues();
    const primaryUrl = uploaded[0].url;
    const extraUrls = uploaded.slice(1).map(function(f) { return f.url; }).join('|');
    let matched = 0;
    names.forEach(function(value, i) {
      if (String(value[0] || '').trim().toLowerCase() === productName.toLowerCase()) {
        const targetRow = AEST.DATA_START_ROW + i;
        sheet.getRange(targetRow, map[AEST.HEADERS.IMAGE]).setValue(primaryUrl);
        if (map[AEST.HEADERS.GALLERY] && extraUrls) {
          sheet.getRange(targetRow, map[AEST.HEADERS.GALLERY]).setValue(extraUrls);
        }
        matched++;
      }
    });
    if (!matched) throw new Error('No se encontraron filas para el modelo ' + productName + '.');
  }

  limpiarCacheAEST();
  return { ok: true, files: uploaded, message: uploaded.length + ' imagen(es) subida(s) correctamente.' };
}

/** Descuenta stock solo cuando el administrador confirma la venta. */
function confirmarVentaFilaActiva() {
  assertAdmin_();
  const ui = SpreadsheetApp.getUi();
  const sheet = getProductSheet_();
  const active = SpreadsheetApp.getActiveRange();
  if (!active || active.getSheet().getSheetId() !== sheet.getSheetId() || active.getRow() < AEST.DATA_START_ROW) {
    ui.alert('Selecciona primero una celda de la fila del producto que vendiste.');
    return;
  }
  const map = getHeaderMap_(sheet);
  const row = active.getRow();
  const product = String(sheet.getRange(row, map[AEST.HEADERS.PRODUCT]).getDisplayValue());
  const size = String(sheet.getRange(row, map[AEST.HEADERS.SIZE]).getDisplayValue());
  const stockCell = sheet.getRange(row, map[AEST.HEADERS.STOCK]);
  const stock = Math.max(0, Number(stockCell.getValue()) || 0);
  const answer = ui.prompt('Confirmar venta', 'Producto: ' + product + '\nTalla: ' + size + '\nStock actual: ' + stock + '\n\n¿Cuántas unidades vendiste?', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const qty = Number(answer.getResponseText());
  if (!Number.isInteger(qty) || qty < 1 || qty > stock) {
    ui.alert('Cantidad inválida. Debe ser un número entero entre 1 y ' + stock + '.');
    return;
  }
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const current = Math.max(0, Number(stockCell.getValue()) || 0);
    if (qty > current) throw new Error('El stock cambió. Stock actual: ' + current);
    stockCell.setValue(current - qty);
    limpiarCacheAEST();
    ui.alert('Venta confirmada. Nuevo stock: ' + (current - qty));
  } finally {
    lock.releaseLock();
  }
}

/** Limpia la caché pública. */
function limpiarCacheAEST() {
  CacheService.getScriptCache().remove('AEST_PUBLIC_CATALOG_V1');
}

/* -------------------- FUNCIONES INTERNAS -------------------- */

function assertAdmin_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('No se pudo verificar tu cuenta. Abre el panel desde Google Sheets con una cuenta que tenga permiso de edición.');
  }
  const file = DriveApp.getFileById(getSS_().getId());
  const allowed = [];
  const owner = file.getOwner();
  if (owner) allowed.push(owner.getEmail().toLowerCase());
  file.getEditors().forEach(function(user) { allowed.push(user.getEmail().toLowerCase()); });
  if (allowed.indexOf(email.toLowerCase()) === -1) {
    throw new Error('Tu cuenta no tiene permiso para administrar AEST. Solicita acceso de editor al archivo de Google Sheets.');
  }
}

function getSS_() {
  return AEST.SPREADSHEET_ID
    ? SpreadsheetApp.openById(AEST.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getProductSheet_() {
  const ss = getSS_();
  if (AEST.PRODUCT_SHEET_NAME) {
    const named = ss.getSheetByName(AEST.PRODUCT_SHEET_NAME);
    if (!named) throw new Error('No existe la pestaña: ' + AEST.PRODUCT_SHEET_NAME);
    return named;
  }
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== AEST.CONFIG_SHEET_NAME) return sheets[i];
  }
  throw new Error('No encuentro la pestaña de productos.');
}

function ensureConfigSheet_() {
  const ss = getSS_();
  let sheet = ss.getSheetByName(AEST.CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(AEST.CONFIG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['DATO', 'VALOR']]);
    const values = [
      ['NOMBRE_TIENDA', 'AEST'],
      ['LOGO_URL', ''],
      ['PORTADA_URL', ''],
      ['WHATSAPP', AEST.WHATSAPP],
      ['COLOR_PRINCIPAL', '#FFFFFF'],
      ['COLOR_FONDO', '#111111'],
      ['PRESENTACION', 'Poleras oversize · Streetwear']
    ];
    sheet.getRange(2, 1, values.length, 2).setValues(values);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 2);
  }
  return sheet;
}

function readConfig_() {
  const sheet = ensureConfigSheet_();
  const values = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) result[key] = String(values[i][1] == null ? '' : values[i][1]).trim();
  }
  return result;
}

function writeConfigValue_(key, value) {
  const sheet = ensureConfigSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function ensureProductColumns_(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(AEST.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0]
    .map(function(v) { return normalizeHeader_(v); });
  if (headers.indexOf(normalizeHeader_(AEST.HEADERS.STOCK)) === -1) {
    sheet.getRange(AEST.HEADER_ROW, sheet.getLastColumn() + 1).setValue(AEST.HEADERS.STOCK);
  }
  if (headers.indexOf(normalizeHeader_(AEST.HEADERS.GALLERY)) === -1) {
    sheet.getRange(AEST.HEADER_ROW, sheet.getLastColumn() + 1).setValue(AEST.HEADERS.GALLERY);
  }
}

function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol) throw new Error('La hoja de productos no tiene encabezados.');
  const headers = sheet.getRange(AEST.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  const map = {};
  headers.forEach(function(header, i) {
    map[normalizeHeader_(header)] = i + 1;
  });
  const result = {};
  Object.keys(map).forEach(function(key) { result[key] = map[key]; });
  return result;
}

function requireHeaders_(map, headers) {
  headers.forEach(function(header) {
    if (!map[normalizeHeader_(header)]) throw new Error('Falta la columna: ' + header);
  });
}

function normalizeHeader_(value) {
  return String(value || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

function ensureFolders_() {
  const root = DriveApp.getFolderById(AEST.DRIVE_ROOT_ID);
  return {
    logo: getOrCreateFolder_(root, AEST.FOLDERS.LOGO),
    cover: getOrCreateFolder_(root, AEST.FOLDERS.COVER),
    products: getOrCreateFolder_(root, AEST.FOLDERS.PRODUCTS)
  };
}

function getOrCreateFolder_(parent, name) {
  const found = parent.getFoldersByName(name);
  return found.hasNext() ? found.next() : parent.createFolder(name);
}

function driveThumbnail_(value, size) {
  if (!value) return '';
  const text = String(value);
  let id = '';
  const patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /^([a-zA-Z0-9_-]{20,})$/
  ];
  for (let i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    if (match) { id = match[1]; break; }
  }
  if (!id) return text;
  return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w' + (size || 600);
}

function parsePrice_(value) {
  if (typeof value === 'number') return value;
  const cleaned = String(value || '').replace(/[^\d.,-]/g, '').replace(',', '.');
  const parsed = Number(cleaned);
  return isFinite(parsed) ? parsed : 0;
}

function normalizePhone_(phone) {
  const digits = String(phone || '').replace(/\D/g, '');
  if (digits.indexOf('51') === 0) return digits;
  return '51' + digits;
}

function sanitizeFileName_(name) {
  return String(name).replace(/[\\/:*?"<>|#%{}~&]/g, '_').slice(0, 100);
}

