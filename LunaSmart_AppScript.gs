/**
 * LUNA SMART — Apps Script Web App
 * Google Sheets ID: 1Dm5fcTs_URmtv8cwUDV6z_LxuGvdpJmf0ZkxszXzuCk
 *
 * INSTRUCCIONES DE DESPLIEGUE:
 * 1. Abre el Google Sheet → Extensiones → Apps Script
 * 2. Pega este código (reemplaza todo)
 * 3. Guarda (Ctrl+S)
 * 4. Implementar → Nueva implementación
 *    - Tipo: Aplicación web
 *    - Ejecutar como: Yo (tu cuenta Google)
 *    - Quién tiene acceso: Cualquier persona
 * 5. Copia la URL generada y actualiza APPS_SCRIPT_URL en el HTML
 */

// ── CONFIGURACIÓN ──────────────────────────────────────────────────────────
const SHEET_ID = '1Dm5fcTs_URmtv8cwUDV6z_LxuGvdpJmf0ZkxszXzuCk';

const PARROT_API_KEY   = 'pk_AFHobF97QSAeAk2LdsmoWYbY0aJNPngk_f343f0db581f4b17b644f101cb58e461';
const PARROT_STORE_UUID = 'd6c9c246-8ff7-44a9-a641-e38793050097';
const PARROT_BASE_URL  = 'https://api.parrotsoftware.io/v2';

// Nombres exactos de las hojas
const HOJAS = {
  INGRESOS:       'INGRESOS',
  ING_DETALLES:   'INGRESOS DETALLES',
  FACTURAS:       'FACTURAS',
  ART_DETALLES:   'ARTICULOS DETALLES',
  CATALOGO:       'Catálogo Maestro',
  COSTO_PROD:     'Costo de Producto',
  CLIENTES:       'DATOS_CLIENTES',
  PROVEEDORES:    'DATOS_PROVEEDORES',
  CATEGORIAS:     'DATOS_CATEGORIASSUBCATEGORIAS',
  INVENTARIO:     'BD_INVENTARIO_GENERAL',
  CONCILIACION:   'CONCILIACION',
  DATA_INGRESOS:  'DATA_INGRESOS',
};

// ── JSON OUTPUT ─────────────────────────────────────────────────────────────
// Apps Script agrega CORS automáticamente en despliegues públicos ("Cualquier persona").
// No se necesita setHeader — de hecho no está disponible en todos los runtimes.
function _json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function _err(msg) {
  return _json({ status: 'error', msg: msg });
}

// ── UTILIDADES ─────────────────────────────────────────────────────────────
function _getSheet(nombre) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(nombre);
  if (!sh) throw new Error('Hoja no encontrada: ' + nombre);
  return sh;
}

/** Lee todos los valores de una hoja como array de arrays, normalizando fechas a DD/MM/YYYY. */
function _leerTodo(nombre) {
  const sh = _getSheet(nombre);
  const vals = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  return vals.map(function(row) {
    return row.map(function(cell) {
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, tz, 'dd/MM/yyyy');
      }
      return cell;
    });
  });
}

/** Devuelve filas como array de arrays (incluyendo encabezado). */
function _hoja(nombre) {
  try {
    return { status: 'ok', data: _leerTodo(nombre) };
  } catch(e) {
    return { status: 'error', msg: e.message, data: [] };
  }
}

/**
 * Devuelve la fila donde escribir el PRÓXIMO registro.
 * Escanea desde el FINAL hacia arriba buscando la última fila con dato
 * real en la columna `col` (base 1), y devuelve esa fila + 1.
 *
 * Ventaja vs buscar la primera vacía: ignora filas con formato/validación
 * que quedan después del último registro real, y siempre escribe
 * inmediatamente después del dato más reciente.
 */
function _siguienteFilaLibre(sh, col) {
  var col0 = col - 1;
  var vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    var cell = String(vals[i][col0]).trim();
    if (cell !== '' && cell !== '0' && cell !== '$0.00') {
      return i + 2; // fila Sheets = i+1, siguiente = i+2
    }
  }
  return 2; // solo hay encabezado
}

/**
 * Escribe datos inmediatamente después del último registro real.
 * Col B (FECHA) es el indicador de fila con dato en todas las hojas principales.
 */
function _escribirFila(sh, datos) {
  var fila = _siguienteFilaLibre(sh, 2);
  sh.getRange(fila, 1, 1, datos.length).setValues([datos]);
}

/** Genera un ID incremental del tipo PREFIX-NNNNN */
function _nextId(hoja, prefix) {
  try {
    const sh = _getSheet(hoja);
    const vals = sh.getDataRange().getValues();
    // Buscar IDs existentes con ese prefijo en col A
    const ids = vals.slice(1)
      .map(function(r) { return String(r[0]); })
      .filter(function(v) { return v.indexOf(prefix) === 0; });
    if (ids.length === 0) return prefix + '-00001';
    const nums = ids.map(function(id) { return parseInt(id.split('-').pop(), 10) || 0; });
    const next = Math.max.apply(null, nums) + 1;
    return prefix + '-' + String(next).padStart(5, '0');
  } catch(e) {
    return prefix + '-' + Date.now();
  }
}

function _fechaHoy() {
  const d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

// ── doGet: lectura ─────────────────────────────────────────────────────────
function doGet(e) {
  const accion = (e.parameter && e.parameter.accion) ? e.parameter.accion : '';

  const map = {
    getINGRESOS:                   HOJAS.INGRESOS,
    getINGRESOSDETALLES:           HOJAS.ING_DETALLES,
    getFACTURAS:                   HOJAS.FACTURAS,
    getARTICULOSDETALLES:          HOJAS.ART_DETALLES,
    getCatalogoMaestro:            HOJAS.CATALOGO,
    getCostodeProducto:            HOJAS.COSTO_PROD,
    getDATOSCLIENTES:              HOJAS.CLIENTES,
    getDATOSPROVEEDORES:           HOJAS.PROVEEDORES,
    getDATASCATEGORIASSUBCATEGORIAS: HOJAS.CATEGORIAS,
    getBDINVENTARIOGENERAL:        HOJAS.INVENTARIO,
    getCONCILIACION:               HOJAS.CONCILIACION,
    getDATAINGRESOS:               HOJAS.DATA_INGRESOS,
  };

  if (accion === 'getUSUARIOS') return _getUsuarios();

  if (map[accion]) {
    return _json(_hoja(map[accion]));
  }

  // Sin acción: devuelve info de salud
  return _json({ status: 'ok', app: 'LunaSmart', version: '4.0', timestamp: new Date().toISOString() });
}

// ── doPost: escritura y sincronización ────────────────────────────────────
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch(_) {
    return _err('JSON inválido');
  }

  const accion = body.accion || '';

  switch (accion) {
    case 'registrarIngreso':        return _registrarIngreso(body);
    case 'registrarFactura':        return _registrarFactura(body);
    case 'registrarArticuloDetalle': return _registrarArticuloDetalle(body);
    case 'registrarProveedor':      return _registrarProveedor(body);
    case 'registrarCliente':        return _registrarCliente(body);
    case 'sincronizarParrot':          return _sincronizarParrot(body.sucursal || 'CASA DE LA CULTURA');
    case 'registrarCatalogoArticulo':  return _registrarCatalogoArticulo(body.datos || body);
    default: return _err('Acción desconocida: ' + accion);
  }
}

// ── REGISTRAR INGRESO ──────────────────────────────────────────────────────
function _registrarIngreso(b) {
  try {
    const sh = _getSheet(HOJAS.INGRESOS);
    const id = _nextId(HOJAS.INGRESOS, 'INGRESOS');
    const efectivo      = parseFloat(b.efectivo      || 0);
    const tarjeta       = parseFloat(b.tarjeta       || 0);
    const transferencia = parseFloat(b.transferencia || 0);
    const rappi         = parseFloat(b.rappi         || 0);
    const total         = efectivo + tarjeta + transferencia + rappi;
    const inicioCaja    = parseFloat(b.inicioCaja    || 0);
    const retiros       = parseFloat(b.retiros       || 0);
    const ventaTotal    = parseFloat(b.ventaTotal    || total);
    const diferencia    = total - ventaTotal;

    _escribirFila(sh, [
      id,
      b.fecha || _fechaHoy(),
      b.sucursal || '',
      b.cliente  || '',          // turno ID
      inicioCaja,
      retiros,
      parseFloat(b.depositos || 0),
      efectivo,
      tarjeta,
      transferencia,
      rappi,
      total,
      ventaTotal,
      diferencia,
      b.observaciones || 'MANUAL',
    ]);

    return _json({ status: 'ok', idIngreso: id });
  } catch(e) {
    return _err(e.message);
  }
}

// ── REGISTRAR FACTURA ──────────────────────────────────────────────────────
// Estructura real del Sheet FACTURAS (7 cols, sin TIPO):
// A=ID_FACTURA | B=FECHA | C=UNIDAD DE NEGOCIO | D=PROVEEDOR |
// E=FOLIO/TICKET | F=FOTO COMPROBANTE | G=TOTAL FACTURA
function _registrarFactura(b) {
  try {
    const sh = _getSheet(HOJAS.FACTURAS);
    const id = _nextId(HOJAS.FACTURAS, 'FACT');

    _escribirFila(sh, [
      id,
      b.fecha     || _fechaHoy(),
      b.unidad    || '',
      b.proveedor || '',
      b.folio     || '',   // col E (sin TIPO intermedio)
      b.foto      || '',
      parseFloat(b.total || 0),
    ]);

    return _json({ status: 'ok', idFactura: id });
  } catch(e) {
    return _err(e.message);
  }
}

// ── REGISTRAR ARTÍCULO DETALLE ─────────────────────────────────────────────
// Estructura real del Sheet ARTICULOS DETALLES (9 cols):
// A=Z(auto) | B=ID_FACTURA | C=FECHA | D=ARTICULO | E=CANTIDAD |
// F=PRECIO UNIT | G=¿APLICA IVA? | H=IVA MONTO | I=TOTAL
// → Escribimos desde col B (dejamos col A para el contador automático del Sheet)
function _registrarArticuloDetalle(b) {
  try {
    const sh = _getSheet(HOJAS.ART_DETALLES);
    const qty   = parseFloat(b.cantidad   || 0);
    const precio= parseFloat(b.precioUnit || 0);
    const sub   = qty * precio;
    const aplica= !!(b.aplicaIva);
    const iva   = aplica ? sub * 0.16 : 0;
    const total = sub + iva;

    // Escribir después del último registro con dato en col B (ID_FACTURA)
    var fila = _siguienteFilaLibre(sh, 2);
    sh.getRange(fila, 2, 1, 8).setValues([[
      b.idFactura  || '',   // B
      b.fecha      || _fechaHoy(), // C
      b.articulo   || '',   // D
      qty,                  // E
      precio,               // F
      aplica ? 'SI' : 'NO', // G
      iva,                  // H
      total,                // I
    ]]);

    // Actualizar Costo Dinámico en Catálogo Maestro
    _actualizarCostoDinamico(b.articulo, precio);

    return _json({ status: 'ok', idDetalle: 'ARTDET-' + fila });
  } catch(e) {
    return _err(e.message);
  }
}

// _siguienteFilaLibre reemplaza a _primeraFilaVaciaDesdeCol en todos los usos.

/** Actualiza la columna Costo Dinámico (col C) en Catálogo Maestro */
function _actualizarCostoDinamico(nombreArticulo, nuevoCosto) {
  if (!nombreArticulo || !nuevoCosto) return;
  try {
    const sh = _getSheet(HOJAS.CATALOGO);
    const datos = sh.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toLowerCase() === String(nombreArticulo).trim().toLowerCase()) {
        sh.getRange(i + 1, 3).setValue(parseFloat(nuevoCosto)); // col C = Costo Dinámico
        break;
      }
    }
  } catch(_) {}
}

// ── REGISTRAR PROVEEDOR ────────────────────────────────────────────────────
function _registrarProveedor(b) {
  try {
    const sh = _getSheet(HOJAS.PROVEEDORES);
    // Evitar duplicados
    const datos = sh.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][1]).trim().toLowerCase() === String(b.nombre || '').trim().toLowerCase()) {
        return _json({ status: 'ok', msg: 'El proveedor ya existe', idProveedor: datos[i][0] });
      }
    }
    const id = _nextId(HOJAS.PROVEEDORES, 'PROV');
    _escribirFila(sh, [
      id,
      b.nombre    || '',
      b.rfc       || '',
      b.contacto  || '',
      b.correo    || '',
      b.telefono  || '',
      b.lineaCredito || 0,
      b.diasCredito  || 0,
    ]);
    return _json({ status: 'ok', idProveedor: id });
  } catch(e) {
    return _err(e.message);
  }
}

// ── REGISTRAR CLIENTE ──────────────────────────────────────────────────────
function _registrarCliente(b) {
  try {
    const sh = _getSheet(HOJAS.CLIENTES);
    const datos = sh.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][1]).trim().toLowerCase() === String(b.nombre || '').trim().toLowerCase()) {
        return _json({ status: 'ok', msg: 'El cliente ya existe', idCliente: datos[i][0] });
      }
    }
    const id = _nextId(HOJAS.CLIENTES, 'INGRESOS');
    _escribirFila(sh, [
      id,
      b.nombre    || '',
      b.rfc       || '',
      b.contacto  || '',
      b.correo    || '',
      b.telefono  || '',
      b.lineaCredito || 0,
      b.diasCredito  || 0,
    ]);
    return _json({ status: 'ok', idCliente: id });
  } catch(e) {
    return _err(e.message);
  }
}

// ── SINCRONIZAR PARROT POS ─────────────────────────────────────────────────
// Sucursales con Parrot POS (UUID compartido):
//   CASA DE LA CULTURA → sucursal principal (punto de venta físico)
//   HELFY FÜ          → multi-marca Parrot (Rappi/Uber)
// Sucursales SIN Parrot (manuales):
//   SUEÑO DE LUNA     → Coffee & Roasters (expos, mercados, redes sociales)
//   BARBACOA Y MENUDO → solo domingos, desde casa
//   EVENTOS           → bar móvil y eventos
function _sincronizarParrot(sucursal) {
  sucursal = sucursal || 'CASA DE LA CULTURA';
  try {
    const hoy = new Date();
    const hace7 = new Date(hoy - 7 * 24 * 60 * 60 * 1000);

    const fechaDesde = Utilities.formatDate(hace7, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
    const fechaHasta = Utilities.formatDate(hoy,   'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");

    // Llamar a la API de Parrot para obtener órdenes
    const url = `${PARROT_BASE_URL}/stores/${PARROT_STORE_UUID}/orders?created_after=${fechaDesde}&created_before=${fechaHasta}&limit=500`;
    const resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + PARROT_API_KEY,
        'Content-Type':  'application/json',
      },
      muteHttpExceptions: true,
    });

    const code = resp.getResponseCode();
    if (code !== 200) {
      // Parrot devuelve error — intentar con endpoint alternativo
      return _sincronizarParrotFallback();
    }

    const json = JSON.parse(resp.getContentText());
    const ordenes = json.results || json.orders || json.data || [];

    if (ordenes.length === 0) {
      return _json({ status: 'ok', msg: 'Sin órdenes nuevas en los últimos 7 días', registros: 0 });
    }

    // Agrupar por fecha y turno
    const turnos = {};
    ordenes.forEach(orden => {
      const fechaRaw = orden.created_at || orden.date || '';
      const fecha = Utilities.formatDate(new Date(fechaRaw), Session.getScriptTimeZone(), 'dd/MM/yyyy');
      const hora  = new Date(fechaRaw).getHours();
      const turno = hora < 13 ? 'INGRESOS-00013' : 'INGRESOS-00015';
      const key   = fecha + '|' + turno;
      if (!turnos[key]) turnos[key] = { fecha, turno, total: 0, count: 0 };
      const monto = parseFloat(orden.total || orden.subtotal || 0);
      turnos[key].total += monto;
      turnos[key].count++;
    });

    // Escribir en CONCILIACION
    const shConc = _getSheet(HOJAS.CONCILIACION);
    const existentes = shConc.getDataRange().getValues().slice(1)
      .map(r => r[0] + '|' + r[1]);

    let registros = 0;
    Object.values(turnos).forEach(t => {
      const key = t.fecha + '|' + t.turno;
      if (!existentes.includes(key)) {
        _escribirFila(shConc, [
          t.fecha,
          t.turno,
          t.total,   // VENTA_PARROT
          0,          // VENTA_TOTAL (se llena con el corte manual)
          0,          // DECLARADO
          0,          // DELTA
          'PENDIENTE',
          'PARROT-AUTO',
          new Date().toISOString(),
        ]);
        registros++;
      }
    });

    // También escribir en INGRESOS como registros de tipo Parrot
    const shIng = _getSheet(HOJAS.INGRESOS);
    const existIngIDs = shIng.getDataRange().getValues().slice(1)
      .map(r => String(r[14])); // col O = OBSERVACIONES

    Object.values(turnos).forEach(t => {
      const sesionId = 'PARROT-' + t.fecha.replace(/\//g, '') + '-' + t.turno;
      if (!existIngIDs.includes(sesionId)) {
        const id = _nextId(HOJAS.INGRESOS, 'INGRESOS');
        _escribirFila(shIng, [
          id, t.fecha,
          sucursal, t.turno,
          0, 0, 0,
          0, t.total, 0, 0,  // efectivo=0, tarjeta=total (aproximado Parrot), transf=0, rappi=0
          t.total, t.total, 0,
          sesionId,
        ]);
      }
    });

    return _json({ status: 'ok', msg: 'Parrot sincronizado', registros });
  } catch(e) {
    return _err('Error Parrot: ' + e.message);
  }
}

/** Fallback: intenta con reportes diarios de Parrot */
function _sincronizarParrotFallback() {
  try {
    const hoy = new Date();
    const fechaStr = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const url = `${PARROT_BASE_URL}/stores/${PARROT_STORE_UUID}/sales-summary?date=${fechaStr}`;
    const resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + PARROT_API_KEY },
      muteHttpExceptions: true,
    });

    const code = resp.getResponseCode();
    if (code !== 200) {
      return _json({ status: 'ok', msg: 'Parrot no disponible (code ' + code + '). Reintenta más tarde.', registros: 0 });
    }

    const json = JSON.parse(resp.getContentText());
    const total = parseFloat(json.total || json.net_sales || 0);

    if (total > 0) {
      const shConc = _getSheet(HOJAS.CONCILIACION);
      const fecha  = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      _escribirFila(shConc, [
        fecha, 'INGRESOS-00013',
        total, 0, 0, 0, 'PENDIENTE', 'PARROT-RESUMEN',
        new Date().toISOString(),
      ]);
    }

    return _json({ status: 'ok', msg: 'Resumen Parrot del día importado', registros: total > 0 ? 1 : 0 });
  } catch(e) {
    return _json({ status: 'ok', msg: 'Parrot fallback falló: ' + e.message, registros: 0 });
  }
}

// ── REGISTRAR ARTÍCULO EN CATÁLOGO MAESTRO ────────────────────────────────
// Columnas: ARTICULO | COSTO BASE | COSTO DINAMICO | CANTIDAD | UNIDA DE MEDIDA
//           % MERMA  | COSTO FINAL | CATEGORIA | SUBCATEGORIA | PROVEEDOR
function _registrarCatalogoArticulo(b) {
  try {
    var sh = _getSheet(HOJAS.CATALOGO);
    var vals = sh.getDataRange().getValues();
    // Verificar duplicado (col A = ARTICULO)
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toLowerCase() === String(b.articulo || '').trim().toLowerCase()) {
        return _json({ status: 'ok', msg: 'El artículo ya existe en el Catálogo Maestro' });
      }
    }
    var costo = parseFloat(b.costoBase || 0);
    var merma = parseFloat(b.merma || 0);
    var costoFinal = merma > 0 ? costo / (1 - merma / 100) : costo;

    var fila = _siguienteFilaLibre(sh, 1); // col A = ARTICULO
    sh.getRange(fila, 1, 1, 10).setValues([[
      b.articulo     || '',
      costo,                 // COSTO BASE
      costo,                 // COSTO DINAMICO (igual al base al inicio)
      1,                     // CANTIDAD
      b.unidad       || '',  // UNIDA DE MEDIDA
      merma,                 // % MERMA
      costoFinal,            // COSTO FINAL
      b.categoria    || '',  // CATEGORIA
      b.subcategoria || '',  // SUBCATEGORIA
      b.proveedor    || '',  // PROVEEDOR
    ]]);
    return _json({ status: 'ok', msg: 'Artículo registrado en Catálogo Maestro' });
  } catch(e) {
    return _err(e.message);
  }
}

// ── AUTENTICACIÓN DESDE USUARIOS_APP ──────────────────────────────────────
// Permite al frontend verificar credenciales contra la hoja USUARIOS_APP
// Llamada: GET ?accion=getUSUARIOS (devuelve lista sin contraseñas)
//          POST { accion:'login', email, password }
// NOTA: por seguridad real implementar OAuth; esto es suficiente para uso interno.
function _getUsuarios() {
  try {
    var sh = _getSheet('USUARIOS_APP');
    var vals = sh.getDataRange().getValues();
    // Devolver solo email, nombre, rol, sucursal (sin contraseñas)
    var users = vals.slice(1).filter(function(r){ return r[0]; }).map(function(r){
      return { email: r[0], nombre: r[1], rol: r[2], sucursal: r[3] };
    });
    return _json({ status: 'ok', data: users });
  } catch(e) { return _err(e.message); }
}

// ── LISTAR TIENDAS / TERMINALES PARROT ────────────────────────────────────
// Ejecuta esta función desde el editor de Apps Script para ver todos los UUIDs.
// 1. Selecciona "listarTiendasParrot" en el menú de funciones
// 2. Haz clic en ▶ Ejecutar
// 3. Lee el resultado en el panel "Registro de ejecución" abajo
function listarTiendasParrot() {
  var endpoints = [
    'https://api.parrotsoftware.io/v2/stores/',
    'https://api.parrotsoftware.io/v1/stores/',
    'https://api.parrotsoftware.io/v2/restaurants/',
    'https://api.parrotsoftware.io/v2/locations/',
  ];
  var headers = { 'Authorization': 'Bearer ' + PARROT_API_KEY, 'Content-Type': 'application/json' };

  for (var i = 0; i < endpoints.length; i++) {
    try {
      var resp = UrlFetchApp.fetch(endpoints[i], { headers: headers, muteHttpExceptions: true });
      var code = resp.getResponseCode();
      var body = resp.getContentText();
      Logger.log('Endpoint: ' + endpoints[i] + ' → HTTP ' + code);
      if (code === 200) {
        var data = JSON.parse(body);
        var items = Array.isArray(data) ? data : (data.results || data.stores || data.data || []);
        Logger.log('✅ Tiendas encontradas: ' + items.length);
        items.forEach(function(s) {
          Logger.log('  UUID: ' + (s.uuid || s.id || '?'));
          Logger.log('  Nombre: ' + (s.name || s.display_name || s.title || '?'));
          Logger.log('  Activa: ' + (s.is_active !== undefined ? s.is_active : s.active || '?'));
          Logger.log('---');
        });
        return; // Si encontró, para aquí
      }
      Logger.log('  Respuesta: ' + body.substring(0, 200));
    } catch(e) {
      Logger.log('Error en ' + endpoints[i] + ': ' + e.message);
    }
  }
  Logger.log('⚠️ No se encontraron tiendas. Verifica la API Key de Parrot.');
  Logger.log('API Key usada: ' + PARROT_API_KEY.substring(0, 20) + '...');
}

// ── REORGANIZAR HOJAS ──────────────────────────────────────────────────────
// Llamar una sola vez desde Apps Script editor: reorganizarHojas()
// NO expuesto como endpoint web por seguridad.
function reorganizarHojas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Orden deseado (22 hojas)
  var orden = [
    // 📊 OPERACIONES DIARIAS
    'INGRESOS',
    'CONCILIACION',
    'FACTURAS',
    'ARTICULOS DETALLES',
    'INGRESOS DETALLES',
    // 🍽️ MENÚ Y COSTOS
    'Catálogo Maestro',
    'Recetas',
    'Costo de Producto',
    'Simulador Cotizador',
    'DATA_INGRESOS',
    // 📦 INVENTARIO
    'BD_INVENTARIO_GENERAL',
    // 👥 MAESTROS
    'DATOS_CLIENTES',
    'DATOS_PROVEEDORES',
    'DATOS_CATEGORIASSUBCATEGORIAS',
    'USUARIOS_APP',
    // 📈 ANÁLISIS Y REPORTES
    'Dashboard Anual',
    'Analisis Financiero 💰',
    'Planificación de Gastos',
    // 🗄️ SISTEMA / COTIZACIONES
    'DATA_COTIZACIONES_MAESTRO',
    'ID_COTIZACION_DETALLE',
    'DATA_APP',
    // 🗂️ ARCHIVO LEGADO
    'ZEGRESOS:VIEJO',
  ];

  for (var i = 0; i < orden.length; i++) {
    var sh = ss.getSheetByName(orden[i]);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
      Utilities.sleep(100); // evitar rate limit
    }
  }

  Logger.log('✅ Hojas reorganizadas correctamente.');
}

// ── EXPONER getUSUARIOS en doGet ───────────────────────────────────────────
// (ya incluido en el map de doGet de arriba — agregar manualmente si se necesita)
