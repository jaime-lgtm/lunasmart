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

// Token compartido que autoriza ESCRITURAS. Debe coincidir con
// CONFIG.WRITE_TOKEN en index.html. Cámbialo cuando quieras (en ambos lados).
const WRITE_TOKEN = 'SDL-luna-w-7Kq9mT2pXr5vB';

const PARROT_API_KEY   = 'pk_AFHobF97QSAeAk2LdsmoWYbY0aJNPngk_f343f0db581f4b17b644f101cb58e461';
const PARROT_STORE_UUID = 'd6c9c246-8ff7-44a9-a641-e38793050097';
const PARROT_BASE_URL  = 'https://api.parrot.rest/external';

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
  VENTAS_PARROT:  'VENTAS_PARROT',        // ventas por artículo (de Parrot)
  CLIENTES_DOM:   'CLIENTES_DOMICILIO',   // clientes de domicilio (POS)
  RECETAS:        'BD_RECETAS',
  RECETAS_DET:    'BD_RECETAS_DETALLES',
  TIEMPOS:        'BD_TIEMPOS',
};

// ── JSON OUTPUT ─────────────────────────────────────────────────────────────
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

function _hoja(nombre) {
  try {
    return { status: 'ok', data: _leerTodo(nombre) };
  } catch(e) {
    return { status: 'error', msg: e.message, data: [] };
  }
}

function _siguienteFilaLibre(sh, col) {
  var col0 = col - 1;
  var vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    var cell = String(vals[i][col0]).trim();
    if (cell !== '' && cell !== '0' && cell !== '$0.00') {
      return i + 2;
    }
  }
  return 2;
}

function _escribirFila(sh, datos) {
  var fila = _siguienteFilaLibre(sh, 2);
  sh.getRange(fila, 1, 1, datos.length).setValues([datos]);
}

function _nextId(hoja, prefix) {
  try {
    const sh = _getSheet(hoja);
    const vals = sh.getDataRange().getValues();
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
    getVENTASPARROT:               HOJAS.VENTAS_PARROT,
    getBD_RECETAS:                 HOJAS.RECETAS,
    getBD_RECETAS_DETALLES:        HOJAS.RECETAS_DET,
    getBD_TIEMPOS:                 HOJAS.TIEMPOS,
  };

  if (accion === 'getUSUARIOS') return _getUsuarios();

  if (map[accion]) {
    return _json(_hoja(map[accion]));
  }

  return _json({ status: 'ok', app: 'LunaSmart', version: '4.0', timestamp: new Date().toISOString() });
}

// ── doPost: escritura ──────────────────────────────────────────────────────
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch(_) {
    return _err('JSON inválido');
  }

  const accion = body.accion || '';

  // ── SEGURIDAD: Verificar token (puede venir en diferentes lugares) ──────
  // El token puede venir en:
  // 1. body.token (forma estándar)
  // 2. body.datos.token (si viene anidado)
  // 3. Fallback: rechazar si no está en ningún lado
  const tokenEnviadoEnBody = body.token || '';
  const tokenEnviadoEnDatos = (body.datos && body.datos.token) || '';
  const tokenRecibido = tokenEnviadoEnBody || tokenEnviadoEnDatos;

  if (tokenRecibido !== WRITE_TOKEN) {
    Logger.log('❌ Intento de acceso NO AUTORIZADO');
    Logger.log('  Token esperado: ' + WRITE_TOKEN);
    Logger.log('  Token recibido: ' + (tokenRecibido || 'NINGUNO'));
    Logger.log('  Acción: ' + accion);
    return _err('No autorizado');
  }

  const datos = body.datos || body;

  switch (accion) {
    case 'registrarIngreso':           return _registrarIngreso(datos);
    case 'registrarIngresoCompleto':   return _registrarIngresoCompleto(datos);
    case 'actualizarIngreso':          return _actualizarIngreso(datos);
    case 'borrarIngreso':              return _borrarIngreso(datos);
    case 'actualizarIngresoParrot':    return _actualizarIngresoParrot(datos);
    case 'borrarIngresoParrot':        return _borrarIngresoParrot(datos);
    case 'registrarFactura':           return _registrarFactura(datos);
    case 'registrarFacturaCompleta':   return _registrarFacturaCompleta(datos);
    case 'actualizarFactura':          return _actualizarFactura(datos);
    case 'borrarFactura':              return _borrarFactura(datos);
    case 'registrarArticuloDetalle':   return _registrarArticuloDetalle(datos);
    case 'registrarProveedor':         return _registrarProveedor(datos);
    case 'actualizarProveedor':        return _actualizarProveedor(datos);
    case 'borrarProveedor':            return _borrarProveedor(datos);
    case 'actualizarCatalogoArticulo': return _actualizarCatalogoArticulo(datos);
    case 'borrarCatalogoArticulo':     return _borrarCatalogoArticulo(datos);
    case 'registrarCliente':           return _registrarCliente(datos);
    case 'registrarReceta':            return _registrarReceta(datos);
    case 'actualizarReceta':           return _actualizarReceta(datos);
    case 'borrarReceta':               return _borrarReceta(datos);
    case 'registrarTiempos':           return _registrarTiempos(datos);
    case 'crearPreferenciaMP':         return _crearPreferenciaMP(datos);
    case 'mpListarTerminales':         return _mpListarTerminales();
    case 'mpCrearCobroTerminal':       return _mpCrearCobroTerminal(datos);
    case 'mpConsultarCobroTerminal':   return _mpConsultarCobroTerminal(datos);
    case 'mpCancelarCobroTerminal':    return _mpCancelarCobroTerminal(datos);
    case 'guardarInventarioFisico':    return _guardarInventarioFisico(datos);
    case 'sumarStockPorCompra':        return _sumarStockPorCompra(datos);
    case 'actualizarMinMaxInventario': return _actualizarMinMaxInventario(datos);
    case 'agregarArticuloInventario':  return _agregarArticuloInventario(datos);
    case 'descontarInventarioVenta':   return _descontarInventarioVenta(datos);
    case 'registrarCatalogoArticulo':  return _registrarCatalogoArticulo(datos);
    case 'buscarClienteDom':           return _buscarClienteDom(datos);
    case 'registrarClienteDom':        return _registrarClienteDom(datos);
    case 'obtenerCorteDia':            return _obtenerCorteDia(datos);
    default: return _err('Acción desconocida: ' + accion);
  }
}

// ── REGISTRAR INGRESO + DETALLE ────────────────────────────────────────────
function _registrarIngresoCompleto(b) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return _err('Sistema ocupado, intenta de nuevo en unos segundos');
  }
  try {
    var sh = _getSheet(HOJAS.INGRESOS);

    // Migración additiva de la columna de idempotencia (ver abajo también las P-W).
    if (!sh.getRange(1, 24).getValue()) { sh.getRange(1, 24).setValue('Corte ID'); }

    // Idempotencia: si este mismo corte (mismo corteId) ya fue registrado
    // -- por ejemplo por un reintento o un doble clic -- no se vuelve a insertar.
    if (b.corteId) {
      var valsCheck = sh.getDataRange().getValues();
      for (var ci = 1; ci < valsCheck.length; ci++) {
        if (String(valsCheck[ci][23]) === String(b.corteId)) {
          return _json({ status: 'ok', idIngreso: valsCheck[ci][0], detalles: 0, duplicado: true });
        }
      }
    }

    var id = _nextId(HOJAS.INGRESOS, 'INGRESOS');
    var efectivo      = parseFloat(b.efectivo      || 0);
    var tarjeta       = parseFloat(b.tarjeta       || 0);
    var mercadopago   = parseFloat(b.mercadopago   || 0);
    var transferencia = parseFloat(b.transferencia || 0);
    var rappi         = parseFloat(b.rappi         || 0);
    var total         = efectivo + tarjeta + mercadopago + transferencia + rappi;
    var inicioCaja    = parseFloat(b.inicioCaja    || 0);
    var retiros       = parseFloat(b.retiros       || 0);
    var ventaTotal    = parseFloat(b.ventaTotal    || total);
    var puntosCanjeados = parseFloat(b.puntosCanjeados || 0);
    var fecha         = b.fecha || _fechaHoy();

    // Migración additiva: la hoja ya existía sin estas columnas -- se agregan
    // los encabezados solo si aún no están, sin tocar las columnas A-O existentes.
    if (!sh.getRange(1, 16).getValue()) { sh.getRange(1, 16).setValue('Puntos Canjeados ($)'); }
    if (!sh.getRange(1, 17).getValue()) { sh.getRange(1, 17).setValue('Mercado Pago ($)'); }
    if (!sh.getRange(1, 18).getValue()) { sh.getRange(1, 18).setValue('Efectivo Declarado ($)'); }
    if (!sh.getRange(1, 19).getValue()) { sh.getRange(1, 19).setValue('Efectivo Diferencia ($)'); }
    if (!sh.getRange(1, 20).getValue()) { sh.getRange(1, 20).setValue('Terminal Declarado ($)'); }
    if (!sh.getRange(1, 21).getValue()) { sh.getRange(1, 21).setValue('Terminal Diferencia ($)'); }
    if (!sh.getRange(1, 22).getValue()) { sh.getRange(1, 22).setValue('Transferencia Declarado ($)'); }
    if (!sh.getRange(1, 23).getValue()) { sh.getRange(1, 23).setValue('Transferencia Diferencia ($)'); }

    _escribirFila(sh, [
      id, fecha, b.sucursal || '', b.cliente || '',
      inicioCaja, retiros, parseFloat(b.depositos || 0),
      efectivo, tarjeta, transferencia, rappi,
      total, ventaTotal, total - ventaTotal,
      b.observaciones || 'MANUAL',
      puntosCanjeados,
      mercadopago,
      (typeof b.cajaDeclarada === 'number') ? b.cajaDeclarada : '',
      (typeof b.cajaDiferencia === 'number') ? b.cajaDiferencia : '',
      (typeof b.terminalDeclarada === 'number') ? b.terminalDeclarada : '',
      (typeof b.terminalDiferencia === 'number') ? b.terminalDiferencia : '',
      (typeof b.transferenciaDeclarada === 'number') ? b.transferenciaDeclarada : '',
      (typeof b.transferenciaDiferencia === 'number') ? b.transferenciaDiferencia : '',
      b.corteId || '',
    ]);

    var nDet = 0;
    var detalles = b.detalles || [];
    if (detalles.length) {
      var shD = _getSheet(HOJAS.ING_DETALLES);
      var filas = detalles.filter(function(d){ return d && d.articulo; }).map(function(d){
        var cant = parseFloat(d.cantidad || 0) || 0;
        var precio = parseFloat(d.precio || 0) || 0;
        return [
          Utilities.getUuid().substring(0, 8),
          id,
          fecha,
          d.articulo,
          cant,
          precio,
          cant * precio,
          '',
          0,
        ];
      });
      if (filas.length) {
        var filaInicio = _siguienteFilaLibre(shD, 4);
        shD.getRange(filaInicio, 1, filas.length, 9).setValues(filas);
        nDet = filas.length;
      }
    }

    return _json({ status: 'ok', idIngreso: id, detalles: nDet });
  } catch (e) {
    return _err(e.message);
  } finally {
    lock.releaseLock();
  }
}

// ── Consultar el corte ya registrado de un turno del día (para acumulado) ──
function _obtenerCorteDia(b) {
  try {
    var sh = _getSheet(HOJAS.INGRESOS);
    var vals = sh.getDataRange().getValues();
    var fecha = b.fecha || _fechaHoy();
    var sucursal = b.sucursal || '';
    var turno = b.turno || 'TURNO MAÑANA';
    for (var i = vals.length - 1; i >= 1; i--) {
      var r = vals[i];
      if (String(r[1]) === String(fecha) && String(r[2]) === String(sucursal) && String(r[3]) === String(turno)) {
        return _json({
          ok: true, encontrado: true,
          fecha: r[1], sucursal: r[2], turno: r[3],
          efectivo: r[7], tarjeta: r[8], transferencia: r[9], rappi: r[10],
          totalDeclarado: r[11], ventaTotal: r[12],
          puntosCanjeados: r[15] || 0, mercadopago: r[16] || 0,
        });
      }
    }
    return _json({ ok: true, encontrado: false });
  } catch (e) {
    return _err('Error al buscar corte del día: ' + e.message);
  }
}

// ── ACTUALIZAR INGRESO ─────────────────────────────────────────────────────
function _actualizarIngreso(b) {
  try {
    var id = String(b.id || '').trim();
    if (!id) return _err('Falta el ID del corte');
    var sh = _getSheet(HOJAS.INGRESOS);
    var vals = sh.getDataRange().getValues();
    var fila = -1;
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === id) { fila = i + 1; break; }
    }
    if (fila === -1) return _err('Corte no encontrado: ' + id);

    var efectivo      = parseFloat(b.efectivo || 0);
    var tarjeta       = parseFloat(b.tarjeta || 0);
    var transferencia = parseFloat(b.transferencia || 0);
    var rappi         = parseFloat(b.rappi || 0);
    var total         = efectivo + tarjeta + transferencia + rappi;
    var ventaTotal    = parseFloat(b.ventaTotal || total);

    sh.getRange(fila, 2, 1, 14).setValues([[
      b.fecha || vals[fila-1][1],
      b.sucursal || '',
      b.cliente || '',
      parseFloat(b.inicioCaja || 0),
      parseFloat(b.retiros || 0),
      parseFloat(b.depositos || 0),
      efectivo, tarjeta, transferencia, rappi,
      total, ventaTotal, total - ventaTotal,
      b.observaciones || 'MANUAL',
    ]]);

    if (b.detalles) {
      var shD = _getSheet(HOJAS.ING_DETALLES);
      var dvals = shD.getDataRange().getValues();
      for (var j = dvals.length - 1; j >= 1; j--) {
        if (String(dvals[j][1]).trim() === id) shD.deleteRow(j + 1);
      }
      var filas = (b.detalles || []).filter(function(d){ return d && d.articulo; }).map(function(d){
        var c = parseFloat(d.cantidad || 0) || 0, p = parseFloat(d.precio || 0) || 0;
        return [Utilities.getUuid().substring(0,8), id, b.fecha || '', d.articulo, c, p, c*p, '', 0];
      });
      if (filas.length) {
        var fi = _siguienteFilaLibre(shD, 4);
        shD.getRange(fi, 1, filas.length, 9).setValues(filas);
      }
    }
    return _json({ status: 'ok', idIngreso: id });
  } catch (e) { return _err(e.message); }
}

// ── BORRAR INGRESO ─────────────────────────────────────────────────────────
function _borrarIngreso(b) {
  try {
    var id = String(b.id || '').trim();
    if (!id) return _err('Falta el ID del corte');
    var sh = _getSheet(HOJAS.INGRESOS);
    var vals = sh.getDataRange().getValues();
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]).trim() === id) sh.deleteRow(i + 1);
    }
    var shD = _getSheet(HOJAS.ING_DETALLES);
    var dvals = shD.getDataRange().getValues();
    for (var j = dvals.length - 1; j >= 1; j--) {
      if (String(dvals[j][1]).trim() === id) shD.deleteRow(j + 1);
    }
    return _json({ status: 'ok' });
  } catch (e) { return _err(e.message); }
}

// Normaliza una fecha (Date real o string dd/MM/yyyy, dd-MM-yyyy, yyyy-MM-dd)
// a 'yyyy-MM-dd' para poder comparar fechas que vienen de columnas con
// formato de número distinto entre hojas (INGRESOS vs INGRESOS DETALLES).
function _fechaAISO(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s;
}

// Localiza la fila de un corte de Parrot (sin ID propio en columna A) usando
// fecha + sucursal + cliente(pseudo-turno) + observaciones exactas como huella
// -- es la unica forma de direccionarlo, ya que estas filas no traen ID.
function _localizarFilaIngresoParrot(sh, b) {
  var vals = sh.getDataRange().getValues();
  var fISO = _fechaAISO(b.fechaOriginal);
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() !== '') continue; // solo filas sin ID (Parrot)
    if (_fechaAISO(vals[i][1]) !== fISO) continue;
    if (String(vals[i][2]).trim() !== String(b.sucursalOriginal || '').trim()) continue;
    if (String(vals[i][3]).trim() !== String(b.clienteOriginal || '').trim()) continue;
    if (String(vals[i][14]).trim() !== String(b.observacionesOriginal || '').trim()) continue;
    return i + 1;
  }
  return -1;
}

// ── ACTUALIZAR / BORRAR INGRESO DE PARROT (sin ID propio) ──────────────────
function _actualizarIngresoParrot(b) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return _err('Sistema ocupado, intenta de nuevo en unos segundos'); }
  try {
    var sh = _getSheet(HOJAS.INGRESOS);
    var fila = _localizarFilaIngresoParrot(sh, b);
    if (fila === -1) return _err('No se encontró el corte original (pudo cambiar desde que se abrió, recarga e intenta de nuevo)');

    var efectivo      = parseFloat(b.efectivo      || 0);
    var tarjeta       = parseFloat(b.tarjeta       || 0);
    var transferencia = parseFloat(b.transferencia || 0);
    var rappi         = parseFloat(b.rappi         || 0);
    var total         = efectivo + tarjeta + transferencia + rappi;
    var ventaTotal    = parseFloat(b.ventaTotal || total);

    sh.getRange(fila, 2, 1, 14).setValues([[
      b.fecha || '', b.sucursal || '', b.cliente || '',
      parseFloat(b.inicioCaja || 0), parseFloat(b.retiros || 0), parseFloat(b.depositos || 0),
      efectivo, tarjeta, transferencia, rappi,
      total, ventaTotal, total - ventaTotal,
      b.observaciones || '',
    ]]);

    if (b.detalles) {
      var shD = _getSheet(HOJAS.ING_DETALLES);
      var dvals = shD.getDataRange().getValues();
      var fISOOrig = _fechaAISO(b.fechaOriginal);
      for (var j = dvals.length - 1; j >= 1; j--) {
        if (String(dvals[j][1]).trim() === String(b.clienteOriginal || '').trim() && _fechaAISO(dvals[j][2]) === fISOOrig) {
          shD.deleteRow(j + 1);
        }
      }
      var filas = (b.detalles || []).filter(function(d){ return d && d.articulo; }).map(function(d){
        var c = parseFloat(d.cantidad || 0) || 0, p = parseFloat(d.precio || 0) || 0;
        return [Utilities.getUuid().substring(0, 8), b.cliente || b.clienteOriginal, b.fecha || b.fechaOriginal, d.articulo, c, p, c * p, '', 0];
      });
      if (filas.length) {
        var fi = _siguienteFilaLibre(shD, 4);
        shD.getRange(fi, 1, filas.length, 9).setValues(filas);
      }
    }
    return _json({ status: 'ok' });
  } catch (e) { return _err(e.message); }
  finally { lock.releaseLock(); }
}

function _borrarIngresoParrot(b) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return _err('Sistema ocupado, intenta de nuevo en unos segundos'); }
  try {
    var sh = _getSheet(HOJAS.INGRESOS);
    var fila = _localizarFilaIngresoParrot(sh, b);
    if (fila === -1) return _err('No se encontró el corte original (pudo cambiar desde que se abrió, recarga e intenta de nuevo)');
    sh.deleteRow(fila);

    var shD = _getSheet(HOJAS.ING_DETALLES);
    var dvals = shD.getDataRange().getValues();
    var fISOOrig = _fechaAISO(b.fechaOriginal);
    for (var j = dvals.length - 1; j >= 1; j--) {
      if (String(dvals[j][1]).trim() === String(b.clienteOriginal || '').trim() && _fechaAISO(dvals[j][2]) === fISOOrig) {
        shD.deleteRow(j + 1);
      }
    }
    return _json({ status: 'ok' });
  } catch (e) { return _err(e.message); }
  finally { lock.releaseLock(); }
}

function asignarIdsIngresos() {
  var sh = _getSheet(HOJAS.INGRESOS);
  var vals = sh.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < vals.length; i++) {
    var m = String(vals[i][0]).match(/INGRESOS-(\d+)/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  var ch = 0;
  for (var k = 1; k < vals.length; k++) {
    if (!String(vals[k][0]).trim() && String(vals[k][1]).trim()) {
      max++;
      sh.getRange(k + 1, 1).setValue('INGRESOS-' + String(max).padStart(5, '0'));
      ch++;
    }
  }
  Logger.log('✅ ' + ch + ' cortes recibieron ID nuevo.');
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
      b.cliente  || '',
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
function _registrarFactura(b) {
  try {
    const sh = _getSheet(HOJAS.FACTURAS);
    const id = _nextId(HOJAS.FACTURAS, 'FACT');

    _escribirFila(sh, [
      id,
      b.fecha     || _fechaHoy(),
      b.unidad    || '',
      b.proveedor || '',
      b.folio     || '',
      b.foto      || '',
      parseFloat(b.total || 0),
    ]);

    return _json({ status: 'ok', idFactura: id });
  } catch(e) {
    return _err(e.message);
  }
}

// ── REGISTRAR FACTURA COMPLETA ────────────────────────────────────────────
function _registrarFacturaCompleta(b) {
  try {
    var shF = _getSheet(HOJAS.FACTURAS);
    var id = _nextId(HOJAS.FACTURAS, 'FACT');
    _escribirFila(shF, [
      id, b.fecha || _fechaHoy(), b.unidad || '', b.proveedor || '',
      b.folio || '', '', parseFloat(b.total || 0)
    ]);

    var lineas = b.lineas || [];
    if (lineas.length > 0) {
      var shD = _getSheet(HOJAS.ART_DETALLES);
      var filaInicio = _siguienteFilaLibre(shD, 2);
      var filas = lineas.map(function(l){
        var qty = parseFloat(l.cantidad) || 0;
        var precio = parseFloat(l.precioUnit) || 0;
        var sub = qty * precio;
        return [ Utilities.getUuid().substring(0,8), id, b.fecha || '', l.articulo || '', qty, precio, 'NO', 0, sub ];
      });
      shD.getRange(filaInicio, 1, filas.length, 9).setValues(filas);
      var mapa = {};
      lineas.forEach(function(l){ if (l.articulo) mapa[String(l.articulo).trim().toLowerCase()] = parseFloat(l.precioUnit) || 0; });
      _actualizarCostosDinamicos(mapa);
    }
    return _json({ status: 'ok', idFactura: id });
  } catch(e) {
    return _err(e.message);
  }
}

function _actualizarCostosDinamicos(mapaArticuloPrecio) {
  try {
    var sh = _getSheet(HOJAS.CATALOGO);
    var datos = sh.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      var nombre = String(datos[i][0]).trim().toLowerCase();
      if (mapaArticuloPrecio.hasOwnProperty(nombre) && mapaArticuloPrecio[nombre] > 0) {
        sh.getRange(i + 1, 3).setValue(mapaArticuloPrecio[nombre]);
      }
    }
  } catch(_) {}
}

// ── ACTUALIZAR FACTURA ────────────────────────────────────────────────────
function _actualizarFactura(b) {
  try {
    var idF = String(b.id || '').trim();
    if (!idF) return _err('Falta el ID de la factura');
    var sh = _getSheet(HOJAS.FACTURAS);
    var vals = sh.getDataRange().getValues();
    var fila = -1;
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === idF) { fila = i + 1; break; }
    }
    if (fila === -1) return _err('Factura no encontrada: ' + idF);

    sh.getRange(fila, 2, 1, 6).setValues([[
      b.fecha     || vals[fila-1][1],
      b.unidad    || '',
      b.proveedor || '',
      b.folio     || '',
      vals[fila-1][5] || '',
      parseFloat(b.total || 0)
    ]]);

    if (b.lineas) {
      var shD = _getSheet(HOJAS.ART_DETALLES);
      var dvals = shD.getDataRange().getValues();
      var aBorrar = [];
      for (var j = 1; j < dvals.length; j++) {
        if (String(dvals[j][1]).trim() === idF) aBorrar.push(j + 1);
      }
      aBorrar.sort(function(a,c){return c-a;}).forEach(function(f){ shD.deleteRow(f); });

      b.lineas.forEach(function(l){
        var qty = parseFloat(l.cantidad) || 0;
        var precio = parseFloat(l.precioUnit) || 0;
        var sub = qty * precio;
        var f2 = _siguienteFilaLibre(shD, 2);
        shD.getRange(f2, 2, 1, 8).setValues([[
          idF, b.fecha || '', l.articulo || '', qty, precio, 'NO', 0, sub
        ]]);
        _actualizarCostoDinamico(l.articulo, precio);
      });
    }
    return _json({ status: 'ok', idFactura: idF });
  } catch(e) {
    return _err(e.message);
  }
}

// ── BORRAR FACTURA ────────────────────────────────────────────────────────
function _borrarFactura(b) {
  try {
    var idF = String(b.id || '').trim();
    if (!idF) return _err('Falta el ID de la factura');
    var sh = _getSheet(HOJAS.FACTURAS);
    var vals = sh.getDataRange().getValues();
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]).trim() === idF) sh.deleteRow(i + 1);
    }
    var shD = _getSheet(HOJAS.ART_DETALLES);
    var dvals = shD.getDataRange().getValues();
    for (var j = dvals.length - 1; j >= 1; j--) {
      if (String(dvals[j][1]).trim() === idF) shD.deleteRow(j + 1);
    }
    return _json({ status: 'ok' });
  } catch(e) {
    return _err(e.message);
  }
}

// ── REGISTRAR ARTÍCULO DETALLE ────────────────────────────────────────────
function _registrarArticuloDetalle(b) {
  try {
    const sh = _getSheet(HOJAS.ART_DETALLES);
    const qty   = parseFloat(b.cantidad   || 0);
    const precio= parseFloat(b.precioUnit || 0);
    const sub   = qty * precio;
    const aplica= !!(b.aplicaIva);
    const iva   = aplica ? sub * 0.16 : 0;
    const total = sub + iva;

    var fila = _siguienteFilaLibre(sh, 2);
    sh.getRange(fila, 2, 1, 8).setValues([[
      b.idFactura  || '',
      b.fecha      || _fechaHoy(),
      b.articulo   || '',
      qty,
      precio,
      aplica ? 'SI' : 'NO',
      iva,
      total,
    ]]);

    _actualizarCostoDinamico(b.articulo, precio);

    return _json({ status: 'ok', idDetalle: 'ARTDET-' + fila });
  } catch(e) {
    return _err(e.message);
  }
}

function _actualizarCostoDinamico(nombreArticulo, nuevoCosto) {
  if (!nombreArticulo || !nuevoCosto) return;
  try {
    const sh = _getSheet(HOJAS.CATALOGO);
    const datos = sh.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toLowerCase() === String(nombreArticulo).trim().toLowerCase()) {
        sh.getRange(i + 1, 3).setValue(parseFloat(nuevoCosto));
        break;
      }
    }
  } catch(_) {}
}

// ── PROVEEDOR ─────────────────────────────────────────────────────────────
function _actualizarProveedor(b) {
  try {
    var id = String(b.id || '').trim();
    if (!id) return _err('Falta el ID del proveedor');
    var sh = _getSheet(HOJAS.PROVEEDORES);
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === id) {
        sh.getRange(i + 1, 2, 1, 7).setValues([[
          b.nombre || '', b.rfc || '', b.contacto || '', b.correo || '',
          b.telefono || '', parseFloat(b.lineaCredito || 0), parseFloat(b.diasCredito || 0),
        ]]);
        return _json({ status: 'ok', idProveedor: id });
      }
    }
    return _err('Proveedor no encontrado: ' + id);
  } catch (e) { return _err(e.message); }
}

function _actualizarCatalogoArticulo(b) {
  try {
    var clave = String(b.articuloKey || b.articulo || '').trim().toLowerCase();
    if (!clave) return _err('Falta el artículo');
    var sh = _getSheet(HOJAS.CATALOGO);
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toLowerCase() === clave) {
        var costo = parseFloat(b.costoBase || vals[i][1] || 0);
        var merma = parseFloat(b.merma || 0); // ej. 30 significa 30% -- se usa asi en la formula de costo
        var costoFinal = merma > 0 ? costo / (1 - merma / 100) : costo;
        var cant = (b.cantidad != null && b.cantidad !== '') ? parseFloat(b.cantidad) : (vals[i][3] || 1);
        sh.getRange(i + 1, 1, 1, 2).setValues([[b.articulo || vals[i][0], costo]]);
        sh.getRange(i + 1, 4, 1, 7).setValues([[
          // La columna % MERMA esta formateada como Porcentaje en el Sheet,
          // que multiplica x100 al mostrar -- por eso se guarda como fraccion
          // (30% -> 0.30), NO como el numero entero que capturo el usuario.
          cant, b.unidad || '', merma / 100, costoFinal,
          b.categoria || '', b.subcategoria || '', b.proveedor || '',
        ]]);
        return _json({ status: 'ok' });
      }
    }
    return _err('Artículo no encontrado: ' + clave);
  } catch (e) { return _err(e.message); }
}

function _borrarCatalogoArticulo(b) {
  try {
    var clave = String(b.articulo || b.articuloKey || '').trim().toLowerCase();
    if (!clave) return _err('Falta el artículo');
    var sh = _getSheet(HOJAS.CATALOGO);
    var vals = sh.getDataRange().getValues();
    var borradas = 0;
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]).trim().toLowerCase() === clave) { sh.deleteRow(i + 1); borradas++; }
    }
    return _json({ status: 'ok', borradas: borradas });
  } catch (e) { return _err(e.message); }
}

function _borrarProveedor(b) {
  try {
    var id = String(b.id || '').trim();
    if (!id) return _err('Falta el ID del proveedor');
    var sh = _getSheet(HOJAS.PROVEEDORES);
    var vals = sh.getDataRange().getValues();
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]).trim() === id) sh.deleteRow(i + 1);
    }
    return _json({ status: 'ok' });
  } catch (e) { return _err(e.message); }
}

function _registrarProveedor(b) {
  try {
    const sh = _getSheet(HOJAS.PROVEEDORES);
    const datos = sh.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][1]).trim().toLowerCase() === String(b.nombre || '').trim().toLowerCase()) {
        return _json({ status: 'ok', msg: 'El proveedor ya existe', idProveedor: datos[i][0] });
      }
    }
    const id = _nextId(HOJAS.PROVEEDORES, 'SL');
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

// ── REGISTRAR RECETA ────────────────────────────────────────────────────────
function _registrarReceta(b) {
  try {
    const sh = _getSheet(HOJAS.RECETAS);
    const id = _nextId(HOJAS.RECETAS, 'REC');
    // Orden real de columnas en BD_RECETAS: ID_RECETA, NOMBRE, CATEGORIA,
    // TIPO, TAMANO, PORCIONES, CLASE, TIPO_VENTA, COSTO_ELABORACION, FECHA.
    _escribirFila(sh, [
      id,
      b.nombre    || '',
      b.categoria || '',
      b.tipo      || '',
      parseFloat(b.tamano    || 0) || 0,
      parseFloat(b.porciones || 1) || 1,
      b.clase || '',
      b.tipoVenta || 'RESTAURANTE',
      parseFloat(b.costoElaboracion || 0) || 0,
      b.fecha || _fechaHoy(),
    ]);

    const ingredientes = b.ingredientes || [];
    if (ingredientes.length) {
      const shD = _getSheet(HOJAS.RECETAS_DET);
      const filas = ingredientes.filter(function(i){ return i && i.articulo; }).map(function(i){
        return [id, i.articulo, parseFloat(i.cantidad || 0) || 0, i.unidad || ''];
      });
      if (filas.length) {
        const fi = _siguienteFilaLibre(shD, 1);
        shD.getRange(fi, 1, filas.length, 4).setValues(filas);
      }
    }
    return _json({ status: 'ok', idReceta: id });
  } catch(e) {
    return _err(e.message);
  }
}

function _actualizarReceta(b) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return _err('Sistema ocupado, intenta de nuevo en unos segundos'); }
  try {
    var id = String(b.id || '').trim();
    if (!id) return _err('Falta el ID de la receta');
    var sh = _getSheet(HOJAS.RECETAS);
    var vals = sh.getDataRange().getValues();
    var fila = -1;
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === id) { fila = i + 1; break; }
    }
    if (fila === -1) return _err('Receta no encontrada: ' + id);

    sh.getRange(fila, 2, 1, 9).setValues([[
      b.nombre    || '',
      b.categoria || '',
      b.tipo      || '',
      parseFloat(b.tamano    || 0) || 0,
      parseFloat(b.porciones || 1) || 1,
      b.clase || '',
      b.tipoVenta || 'RESTAURANTE',
      parseFloat(b.costoElaboracion || 0) || 0,
      b.fecha || vals[fila - 1][9],
    ]]);

    if (b.ingredientes) {
      var shD = _getSheet(HOJAS.RECETAS_DET);
      var dvals = shD.getDataRange().getValues();
      for (var j = dvals.length - 1; j >= 1; j--) {
        if (String(dvals[j][0]).trim() === id) shD.deleteRow(j + 1);
      }
      var filas = (b.ingredientes || []).filter(function(i){ return i && i.articulo; }).map(function(i){
        return [id, i.articulo, parseFloat(i.cantidad || 0) || 0, i.unidad || ''];
      });
      if (filas.length) {
        var fi = _siguienteFilaLibre(shD, 1);
        shD.getRange(fi, 1, filas.length, 4).setValues(filas);
      }
    }
    return _json({ status: 'ok', idReceta: id });
  } catch (e) { return _err(e.message); }
  finally { lock.releaseLock(); }
}

function _borrarReceta(b) {
  try {
    var id = String(b.id || '').trim();
    if (!id) return _err('Falta el ID de la receta');
    var sh = _getSheet(HOJAS.RECETAS);
    var vals = sh.getDataRange().getValues();
    var borrada = false;
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]).trim() === id) { sh.deleteRow(i + 1); borrada = true; break; }
    }
    // Cascada: tambien se borran las lineas de ingredientes de esta receta,
    // si no quedan huerfanas en BD_RECETAS_DETALLES apuntando a un ID que ya no existe.
    var shD = _getSheet(HOJAS.RECETAS_DET);
    var dvals = shD.getDataRange().getValues();
    for (var j = dvals.length - 1; j >= 1; j--) {
      if (String(dvals[j][0]).trim() === id) shD.deleteRow(j + 1);
    }
    if (!borrada) return _err('Receta no encontrada: ' + id);
    return _json({ status: 'ok' });
  } catch (e) { return _err(e.message); }
}

// Un renglon por articulo marcado "listo" en Estacion, capturado justo antes
// de cerrar turno (pedidos/{suc} se borra al cerrar, asi que esto es lo unico
// que sobrevive para poder comparar dias/semanas despues). Columnas de
// BD_TIEMPOS: ID, SUCURSAL, TURNO, FECHA, HORA, PRODUCTO, AREA,
// MINUTOS_ELABORACION, CORTE_ID. El semaforo (verde/amarillo/rojo) NO se
// guarda aqui -- se calcula al vuelo en el reporte usando las metas
// configuradas en ese momento (metasTiempo en Firebase), para que cambiar
// las metas despues recalcule tambien el historial de forma consistente.
function _registrarTiempos(b) {
  try {
    const registros = b.registros || [];
    if (!registros.length) return _json({ status: 'ok', insertados: 0 });
    const sh = _getSheet(HOJAS.TIEMPOS);
    const filas = registros.map(function(r) {
      return [
        Utilities.getUuid().substring(0, 8),
        r.sucursal || '',
        r.turno || '',
        r.fecha || '',
        r.hora || '',
        r.producto || '',
        r.area || '',
        r.minutos != null ? r.minutos : '',
        r.corteId || '',
      ];
    });
    const fi = _siguienteFilaLibre(sh, 1);
    // HORA (columna 5) se guarda como texto plano "HH:MM" -- si no se fuerza
    // el formato ANTES de escribir, Sheets la auto-convierte a un serial de
    // fecha/hora y la columna termina mostrando "30/12/1899" (su fecha cero),
    // perdiendo la hora real que necesita el desglose "Por hora del día".
    sh.getRange(fi, 5, filas.length, 1).setNumberFormat('@');
    sh.getRange(fi, 1, filas.length, 9).setValues(filas);
    return _json({ status: 'ok', insertados: filas.length });
  } catch (e) { return _err(e.message); }
}

// ── MERCADO PAGO (Checkout Pro para pedidos.suenodeluna.com.mx) ───────────
// El Access Token NUNCA va aqui en el codigo (a diferencia de WRITE_TOKEN):
// se puede mover dinero real, asi que se guarda en Propiedades del Script
// (Configuracion del proyecto -> Propiedades del script -> MP_ACCESS_TOKEN),
// que no viaja en el archivo .gs ni queda visible para nadie que solo tenga
// este texto.
function _crearPreferenciaMP(b) {
  try {
    const accessToken = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN');
    if (!accessToken) return _err('Falta configurar MP_ACCESS_TOKEN en Propiedades del script (Configuración del proyecto de Apps Script)');
    const items = (b.items || []).map(function(it) {
      return {
        title: String(it.nombre || '').substring(0, 256) || 'Producto',
        quantity: parseInt(it.cantidad, 10) || 1,
        unit_price: parseFloat(it.precio) || 0,
        currency_id: 'MXN',
      };
    }).filter(function(it) { return it.unit_price > 0; });
    if (!items.length) return _err('Sin artículos válidos para cobrar');
    const baseUrl = b.baseUrl || 'https://pedidos.suenodeluna.com.mx';
    const payload = {
      items: items,
      payer: { name: b.nombre || '', phone: { number: b.telefono || '' } },
      back_urls: {
        success: baseUrl + '?mp_status=success',
        failure: baseUrl + '?mp_status=failure',
        pending: baseUrl + '?mp_status=pending',
      },
      auto_return: 'approved',
      external_reference: b.referencia || ('pedido_' + Date.now()),
      statement_descriptor: 'SUENO DE LUNA',
    };
    const resp = UrlFetchApp.fetch('https://api.mercadopago.com/checkout/preferences', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + accessToken },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const data = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() >= 400) {
      return _err('Mercado Pago rechazó la solicitud: ' + (data.message || resp.getContentText()));
    }
    return _json({ status: 'ok', initPoint: data.init_point, preferenceId: data.id });
  } catch (e) { return _err(e.message); }
}

// ── MERCADO PAGO POINT (terminal física, cobro con tarjeta desde el POS) ──
// Requiere que la terminal ya esté dada de alta en modo PDV (integrado)
// desde la app móvil de Mercado Pago: crear tienda + caja, asociar la
// terminal a esa caja, y usar _mpListarTerminales para obtener su
// terminal_id real (algo como "NEWLAND_N950__N950NCB801293324"). Ese id se
// guarda en Firebase (terminalesMP/{sucursal}) desde el admin.
//
// Usa un Access Token SEPARADO de MP_ACCESS_TOKEN (el de Checkout Pro /
// pagos online) porque en Mercado Pago cada aplicación se declara para un
// solo "producto integrado" (Pagos online O Pagos presencial) -- se creó una
// segunda aplicación exclusiva para Point, con su propio token, en vez de
// arriesgar la que ya funciona para pedidos.suenodeluna.com.mx.
function _mpHeaders() {
  var token = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN_POINT');
  if (!token) throw new Error('Falta configurar MP_ACCESS_TOKEN_POINT en Propiedades del script (token de la app de Pagos presenciales, distinto al de pagos en línea)');
  return { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' };
}

function _mpListarTerminales() {
  try {
    var resp = UrlFetchApp.fetch('https://api.mercadopago.com/terminals/v1/list?limit=50&offset=0', {
      method: 'get', headers: _mpHeaders(), muteHttpExceptions: true,
    });
    var data = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() >= 400) return _err('Mercado Pago: ' + (data.message || resp.getContentText()));
    var terminales = data.terminals || data.results || (Array.isArray(data) ? data : []);
    return _json({ status: 'ok', terminales: terminales, _raw: data });
  } catch (e) { return _err(e.message); }
}

function _mpCrearCobroTerminal(b) {
  try {
    var terminalId = b.terminalId;
    var monto = parseFloat(b.monto);
    if (!terminalId) return _err('Falta el ID de la terminal (configúralo en el admin)');
    if (!(monto > 0)) return _err('Monto inválido');
    var payload = {
      type: 'point',
      external_reference: b.referencia || ('pos_' + Date.now()),
      transactions: { payments: [{ amount: monto.toFixed(2) }] },
      config: { point: { terminal_id: terminalId } },
    };
    var headers = _mpHeaders();
    headers['X-Idempotency-Key'] = Utilities.getUuid();
    var resp = UrlFetchApp.fetch('https://api.mercadopago.com/v1/orders', {
      method: 'post', headers: headers, payload: JSON.stringify(payload), muteHttpExceptions: true,
    });
    var data = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() >= 400) return _err('Mercado Pago rechazó el cobro: ' + (data.message || resp.getContentText()));
    return _json({ status: 'ok', orderId: data.id, estado: data.status });
  } catch (e) { return _err(e.message); }
}

function _mpConsultarCobroTerminal(b) {
  try {
    var orderId = b.orderId;
    if (!orderId) return _err('Falta el ID de la orden');
    var resp = UrlFetchApp.fetch('https://api.mercadopago.com/v1/orders/' + orderId, {
      method: 'get', headers: _mpHeaders(), muteHttpExceptions: true,
    });
    var data = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() >= 400) return _err('Mercado Pago: ' + (data.message || resp.getContentText()));
    return _json({ status: 'ok', estado: data.status, detalle: data.status_detail || '' });
  } catch (e) { return _err(e.message); }
}

function _mpCancelarCobroTerminal(b) {
  try {
    var orderId = b.orderId;
    if (!orderId) return _err('Falta el ID de la orden');
    var headers = _mpHeaders();
    headers['X-Idempotency-Key'] = Utilities.getUuid();
    var resp = UrlFetchApp.fetch('https://api.mercadopago.com/v1/orders/' + orderId + '/cancel', {
      method: 'post', headers: headers, muteHttpExceptions: true,
    });
    var data = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() >= 400) return _err('No se pudo cancelar en la terminal: ' + (data.message || resp.getContentText()));
    return _json({ status: 'ok', estado: data.status });
  } catch (e) { return _err(e.message); }
}

// ── INVENTARIO ───────────────────────────────────────────────────────────
// Columnas de BD_INVENTARIO_GENERAL (1-based):
// 1 ID, 2 SUCURSAL, 3 CONTEO_VISUAL, 4 CATEGORIA, 5 UBICACION, 6 ARTICULO,
// 7 STOCK_FISICO, 8 FALTANTE, 9 STOCK_SUGERIDO, 10 UNIDAD, 11 PRESENTACION,
// 12 COSTO_UNIT, 13 CONVERSION, 14 COSTO_RESURTIDO, 15 TIMESTAMP,
// 16 STOCK_MINIMO, 17 TIPO_REGISTRO.

// Guarda un conteo físico -- cada item ya trae la fila exacta (viene de
// getInventario(), que expone `fila` por cada renglón leído).
function _guardarInventarioFisico(b) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return _err('Sistema ocupado, intenta de nuevo en unos segundos'); }
  try {
    var sh = _getSheet(HOJAS.INVENTARIO);
    var items = b.items || [];
    var actualizados = 0;
    items.forEach(function(it) {
      var fila = parseInt(it.fila, 10);
      if (!fila || fila < 2) return;
      sh.getRange(fila, 7).setValue(parseFloat(it.stockFisico || 0) || 0);
      sh.getRange(fila, 15).setValue(new Date());
      actualizados++;
    });
    return _json({ status: 'ok', actualizados: actualizados });
  } catch (e) { return _err(e.message); }
  finally { lock.releaseLock(); }
}

// Suma cantidades compradas (de una factura) al stock físico -- se localiza
// por nombre exacto de artículo porque las líneas de factura no traen fila.
function _sumarStockPorCompra(b) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return _err('Sistema ocupado, intenta de nuevo en unos segundos'); }
  try {
    var sh = _getSheet(HOJAS.INVENTARIO);
    var vals = sh.getDataRange().getValues();
    var items = b.items || [];
    var actualizados = 0;
    items.forEach(function(it) {
      var nombre = String(it.articulo || '').trim().toLowerCase();
      var cant = parseFloat(it.cantidad || 0) || 0;
      if (!nombre || !cant) return;
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][5]).trim().toLowerCase() === nombre) {
          var filaSheet = i + 1;
          var actual = parseFloat(vals[i][6]) || 0;
          var nuevo = actual + cant;
          sh.getRange(filaSheet, 7).setValue(nuevo);
          sh.getRange(filaSheet, 15).setValue(new Date());
          vals[i][6] = nuevo;
          actualizados++;
        }
      }
    });
    return _json({ status: 'ok', actualizados: actualizados });
  } catch (e) { return _err(e.message); }
  finally { lock.releaseLock(); }
}

// Actualiza mínimo/máximo/tipo de registro -- cada item trae fila exacta.
function _actualizarMinMaxInventario(b) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return _err('Sistema ocupado, intenta de nuevo en unos segundos'); }
  try {
    var sh = _getSheet(HOJAS.INVENTARIO);
    var items = b.items || [];
    var actualizados = 0;
    items.forEach(function(it) {
      var fila = parseInt(it.fila, 10);
      if (!fila || fila < 2) return;
      if (it.minimo !== undefined && it.minimo !== '') sh.getRange(fila, 16).setValue(parseFloat(it.minimo) || 0);
      if (it.maximo !== undefined && it.maximo !== '') sh.getRange(fila, 9).setValue(parseFloat(it.maximo) || 0);
      if (it.tipoRegistro) sh.getRange(fila, 17).setValue(it.tipoRegistro);
      actualizados++;
    });
    return _json({ status: 'ok', actualizados: actualizados });
  } catch (e) { return _err(e.message); }
  finally { lock.releaseLock(); }
}

// Agrega un artículo nuevo a inventario (para una sucursal). Si el artículo
// existe en el Catálogo Maestro, hereda categoría/unidad/costo de ahí.
function _agregarArticuloInventario(b) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return _err('Sistema ocupado, intenta de nuevo en unos segundos'); }
  try {
    var articulo = String(b.articulo || '').trim();
    if (!articulo) return _err('Falta el artículo');
    var sucursal = String(b.sucursal || '').trim();
    if (!sucursal) return _err('Falta la sucursal');

    var categoria = '', unidad = '', costoUnit = 0;
    try {
      var shCat = _getSheet(HOJAS.CATALOGO);
      var catVals = shCat.getDataRange().getValues();
      for (var i = 1; i < catVals.length; i++) {
        if (String(catVals[i][0]).trim().toLowerCase() === articulo.toLowerCase()) {
          unidad = catVals[i][4] || '';
          costoUnit = parseFloat(catVals[i][6]) || parseFloat(catVals[i][1]) || 0;
          categoria = catVals[i][7] || '';
          break;
        }
      }
    } catch (e2) {}

    var sh = _getSheet(HOJAS.INVENTARIO);
    var id = _nextId(HOJAS.INVENTARIO, 'INV');
    _escribirFila(sh, [
      id, sucursal, '', categoria, b.ubicacion || '', articulo,
      parseFloat(b.stockInicial || 0) || 0, '', parseFloat(b.maximo || 0) || 0,
      unidad, '', costoUnit, 1, 0, new Date(),
      parseFloat(b.minimo || 0) || 0, b.tipoRegistro || '',
    ]);
    return _json({ status: 'ok', id: id });
  } catch (e) { return _err(e.message); }
  finally { lock.releaseLock(); }
}

// Descuenta stock por una venta del POS -- se localiza por sucursal + nombre
// exacto de articulo (las lineas de venta no traen fila). Si el resultado
// queda por debajo de STOCK_MINIMO se reporta como alerta al POS.
function _descontarInventarioVenta(b) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return _err('Sistema ocupado, intenta de nuevo en unos segundos'); }
  try {
    var sucursal = String(b.sucursal || '').trim().toLowerCase();
    var items = b.items || [];
    if (!sucursal) return _err('Falta la sucursal');
    var sh = _getSheet(HOJAS.INVENTARIO);
    var vals = sh.getDataRange().getValues();
    var actualizados = 0, alertas = [], noEncontrados = [];
    items.forEach(function(it) {
      var nombre = String(it.articulo || '').trim().toLowerCase();
      var cant = parseFloat(it.cantidad || 0) || 0;
      if (!nombre || !cant) return;
      var encontrado = false;
      for (var i = 1; i < vals.length; i++) {
        if (String(vals[i][1]).trim().toLowerCase() === sucursal && String(vals[i][5]).trim().toLowerCase() === nombre) {
          encontrado = true;
          var filaSheet = i + 1;
          var actual = parseFloat(vals[i][6]) || 0;
          var nuevo = actual - cant;
          var minimo = parseFloat(vals[i][15]) || 0;
          sh.getRange(filaSheet, 7).setValue(nuevo);
          sh.getRange(filaSheet, 15).setValue(new Date());
          vals[i][6] = nuevo;
          actualizados++;
          if (minimo > 0 && nuevo < minimo) alertas.push(String(vals[i][5]) + ' (' + nuevo + ')');
        }
      }
      if (!encontrado) noEncontrados.push(it.articulo);
    });
    return _json({ status: 'ok', actualizados: actualizados, alertas: alertas, noEncontrados: noEncontrados });
  } catch (e) { return _err(e.message); }
  finally { lock.releaseLock(); }
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

// ── CLIENTES DE DOMICILIO (POS) ────────────────────────────────────────────
// Hoja aparte, exclusiva del POS. No toca ninguna hoja usada por Luna Smart admin.
function _getOrCrearSheetClientesDom() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(HOJAS.CLIENTES_DOM);
  if (!sh) {
    sh = ss.insertSheet(HOJAS.CLIENTES_DOM);
    sh.appendRow(['Teléfono', 'Nombre', 'Dirección', 'Referencia', 'Sucursal', 'Última actualización', 'Cumpleaños', 'Puntos', 'Etiqueta']);
  } else if (sh.getLastColumn() < 9) {
    // Migración additiva: la hoja ya existía con el esquema viejo (6 columnas).
    sh.getRange(1, 7, 1, 3).setValues([['Cumpleaños', 'Puntos', 'Etiqueta']]);
  }
  return sh;
}

function _normTel(tel) {
  return String(tel || '').replace(/\D/g, '');
}

function _buscarClienteDom(b) {
  try {
    var tel = _normTel(b.telefono);
    if (!tel) return _json({ status: 'ok', ok: false });
    var sh = _getOrCrearSheetClientesDom();
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (_normTel(vals[i][0]) === tel) {
        return _json({
          status: 'ok', ok: true,
          cliente: {
            nombre: vals[i][1] || '', direccion: vals[i][2] || '', referencia: vals[i][3] || '',
            cumpleanos: vals[i][6] || '', puntos: vals[i][7] || 0, etiqueta: vals[i][8] || '',
          }
        });
      }
    }
    return _json({ status: 'ok', ok: false });
  } catch(e) {
    return _err(e.message);
  }
}

function _registrarClienteDom(b) {
  try {
    var tel = _normTel(b.telefono);
    if (!tel) return _err('Falta el teléfono');
    var sh = _getOrCrearSheetClientesDom();
    var vals = sh.getDataRange().getValues();
    var fecha = _fechaHoy();
    for (var i = 1; i < vals.length; i++) {
      if (_normTel(vals[i][0]) === tel) {
        sh.getRange(i + 1, 1, 1, 9).setValues([[
          b.telefono || vals[i][0], b.nombre || vals[i][1], b.direccion || vals[i][2],
          b.referencia || vals[i][3], b.sucursal || vals[i][4] || '', fecha,
          b.cumpleanos || vals[i][6] || '', (b.puntos != null ? b.puntos : vals[i][7]) || 0,
          b.etiqueta || vals[i][8] || '',
        ]]);
        return _json({ status: 'ok' });
      }
    }
    sh.appendRow([
      b.telefono || tel, b.nombre || '', b.direccion || '', b.referencia || '', b.sucursal || '', fecha,
      b.cumpleanos || '', b.puntos || 0, b.etiqueta || '',
    ]);
    return _json({ status: 'ok' });
  } catch(e) {
    return _err(e.message);
  }
}

// ── REGISTRAR ARTÍCULO EN CATÁLOGO MAESTRO ────────────────────────────────
function _registrarCatalogoArticulo(b) {
  try {
    var sh = _getSheet(HOJAS.CATALOGO);
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toLowerCase() === String(b.articulo || '').trim().toLowerCase()) {
        return _json({ status: 'ok', msg: 'El artículo ya existe en el Catálogo Maestro' });
      }
    }
    var costo = parseFloat(b.costoBase || 0);
    var merma = parseFloat(b.merma || 0);
    var costoFinal = merma > 0 ? costo / (1 - merma / 100) : costo;

    var fila = _siguienteFilaLibre(sh, 1);
    sh.getRange(fila, 1, 1, 10).setValues([[
      b.articulo     || '',
      costo,
      costo,
      (b.cantidad != null && b.cantidad !== '') ? parseFloat(b.cantidad) : 1,
      b.unidad       || '',
      merma / 100, // la columna % MERMA esta formateada como Porcentaje: se guarda como fraccion
      costoFinal,
      b.categoria    || '',
      b.subcategoria || '',
      b.proveedor    || '',
    ]]);
    return _json({ status: 'ok', msg: 'Artículo registrado en Catálogo Maestro' });
  } catch(e) {
    return _err(e.message);
  }
}

// ── AUTENTICACIÓN ──────────────────────────────────────────────────────────
function _getUsuarios() {
  try {
    var sh = _getSheet('USUARIOS_APP');
    var vals = sh.getDataRange().getValues();
    var users = vals.slice(1).filter(function(r){ return r[0]; }).map(function(r){
      return { email: r[0], nombre: r[1], rol: r[2], sucursal: r[3] };
    });
    return _json({ status: 'ok', data: users });
  } catch(e) { return _err(e.message); }
}

// ── CATEGORIZACIÓN ────────────────────────────────────────────────────────
function _categoriasUnicas() {
  var sh = _getSheet(HOJAS.CATEGORIAS);
  var vals = sh.getDataRange().getValues();
  var set = {};
  for (var i = 1; i < vals.length; i++) {
    var c = String(vals[i][1] || '').trim();
    if (c) set[c] = true;
  }
  return Object.keys(set).sort();
}

function _subcategoriasDe(categoria) {
  var sh = _getSheet(HOJAS.CATEGORIAS);
  var vals = sh.getDataRange().getValues();
  var subs = [];
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][1] || '').trim() === categoria) {
      var s = String(vals[i][2] || '').trim();
      if (s) subs.push(s);
    }
  }
  return subs;
}

// ── TRIGGER onEdit ────────────────────────────────────────────────────────
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== HOJAS.CATALOGO) return;
    var col = e.range.getColumn();
    var row = e.range.getRow();
    if (col !== 8 || row < 2) return;

    var categoria = String(e.range.getValue() || '').trim();
    var celdaSub = sh.getRange(row, 9);
    if (!categoria) { celdaSub.clearDataValidations(); return; }

    var subs = _subcategoriasDe(categoria);
    if (subs.length > 0) {
      var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(subs, true)
        .setAllowInvalid(false)
        .build();
      celdaSub.setDataValidation(regla);
      celdaSub.clearContent();
    } else {
      celdaSub.clearDataValidations();
    }
  } catch(err) {
    // Silencioso
  }
}

// ── FUNCIÓN PARA VER TODOS LOS CORTES DE COFFEE & ROASTERS ────────────────
function listarCoffeRoasters() {
  const sh = SpreadsheetApp.openById('1Dm5fcTs_URmtv8cwUDV6z_LxuGvdpJmf0ZkxszXzuCk').getSheetByName('INGRESOS');
  const datos = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  const formatearFecha = function(celda) {
    if (celda instanceof Date) {
      return Utilities.formatDate(celda, tz, 'dd-MM-yyyy');
    }
    return String(celda || '').trim();
  };

  Logger.log('=== TODOS LOS CORTES DE COFFEE & ROASTERS ===');
  Logger.log('');

  let encontrados = 0;
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const idFila = String(fila[0] || '').trim();
    const fechaFila = formatearFecha(fila[1]);
    const sucursalFila = String(fila[2] || '').trim().toUpperCase();
    const turnoFila = String(fila[3] || '').trim();

    if (sucursalFila.includes('COFFEE') || sucursalFila.includes('ROASTER')) {
      encontrados++;
      Logger.log('Fila ' + (i+1) + ': ' + idFila + ' | ' + fechaFila + ' | ' + sucursalFila + ' | ' + turnoFila);
    }
  }

  Logger.log('');
  Logger.log('Total encontrados: ' + encontrados);
  if (encontrados === 0) {
    Logger.log('No hay cortes de Coffee & Roasters en el Sheet.');
    Logger.log('Nota: En Luna Smart ve "Coffee & Roasters" pero en el Sheet podría estar bajo otro nombre.');
  }
}

// ── VERIFICAR QUE SE BORRARON LOS CORTES ────────────────────────────────────
function verificarBorrado() {
  const sh = SpreadsheetApp.openById('1Dm5fcTs_URmtv8cwUDV6z_LxuGvdpJmf0ZkxszXzuCk').getSheetByName('INGRESOS');
  const datos = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  const formatearFecha = function(celda) {
    if (celda instanceof Date) {
      return Utilities.formatDate(celda, tz, 'dd-MM-yyyy');
    }
    return String(celda || '').trim();
  };

  Logger.log('=== VERIFICANDO: ¿SE BORRARON LAS 4 FILAS? ===');
  Logger.log('');
  Logger.log('Buscando las filas que DEBERÍAN estar borradas:');
  Logger.log('  20-06-2026 | CASA DE LA CULTURA | TURNO MAÑANA & TARDE');
  Logger.log('  22-06-2026 | CASA DE LA CULTURA | TURNO MAÑANA');
  Logger.log('  23-06-2026 | CASA DE LA CULTURA | TURNO MAÑANA');
  Logger.log('  23-06-2026 | CASA DE LA CULTURA | TURNO TARDE');
  Logger.log('');

  let encontrados = 0;
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const fechaFila = formatearFecha(fila[1]);
    const sucursalFila = String(fila[2] || '').trim();
    const turnoFila = String(fila[3] || '').trim();

    if ((fechaFila === '20-06-2026' || fechaFila === '22-06-2026' || fechaFila === '23-06-2026') &&
        sucursalFila.includes('CULTURA')) {
      encontrados++;
      Logger.log('❌ ENCONTRADO (DEBERÍA ESTAR BORRADO): Fila ' + (i+1) + ' | ' + fechaFila + ' | ' + sucursalFila + ' | ' + turnoFila);
    }
  }

  Logger.log('');
  if (encontrados === 0) {
    Logger.log('✅ CONFIRMADO: Las 4 filas fueron BORRADAS CORRECTAMENTE del Sheet.');
    Logger.log('');
    Logger.log('Si aún ves datos en Luna Smart, es que:');
    Logger.log('  1. Luna Smart está cacheando en el navegador');
    Logger.log('  2. O lee de otra hoja diferente a "INGRESOS"');
    Logger.log('  3. O hay un delay de sincronización');
    Logger.log('');
    Logger.log('SOLUCIÓN: Presiona Ctrl+F5 en Luna Smart para limpiar caché');
  } else {
    Logger.log('⚠️ WARNING: Encontré ' + encontrados + ' fila(s) que NO deberían estar ahí.');
  }
}

// ── LISTAR TODOS LOS CORTES DE JUNIO 2026 ────────────────────────────────
function listarJunio2026() {
  const sh = SpreadsheetApp.openById('1Dm5fcTs_URmtv8cwUDV6z_LxuGvdpJmf0ZkxszXzuCk').getSheetByName('INGRESOS');
  const datos = sh.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  const formatearFecha = function(celda) {
    if (celda instanceof Date) {
      return Utilities.formatDate(celda, tz, 'dd-MM-yyyy');
    }
    return String(celda || '').trim();
  };

  Logger.log('=== CORTES DE JUNIO 2026 (TODOS) ===');
  Logger.log('');

  let encontrados = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const idFila = String(fila[0] || '').trim();
    const fechaFila = formatearFecha(fila[1]);
    const sucursalFila = String(fila[2] || '').trim();
    const turnoFila = String(fila[3] || '').trim();

    // Buscar fechas que contengan "06-" (junio)
    if (fechaFila.includes('-06-2026')) {
      encontrados.push({
        fila: i + 1,
        id: idFila,
        fecha: fechaFila,
        sucursal: sucursalFila,
        turno: turnoFila
      });
      Logger.log('Fila ' + (i+1) + ': ' + idFila + ' | ' + fechaFila + ' | ' + sucursalFila + ' | ' + turnoFila);
    }
  }

  Logger.log('');
  Logger.log('═══════════════════════════════════════════');
  Logger.log('TOTAL ENCONTRADOS EN JUNIO 2026: ' + encontrados.length);
  Logger.log('═══════════════════════════════════════════');

  if (encontrados.length === 0) {
    Logger.log('⚠️ No hay cortes de junio en el Sheet.');
  }
}

// ── FUNCIÓN PARA BORRAR CORTES EQUIVOCADOS (VERSIÓN MEJORADA) ─────────────
function borrarCortesEquivocados() {
  const ss = SpreadsheetApp.openById('1Dm5fcTs_URmtv8cwUDV6z_LxuGvdpJmf0ZkxszXzuCk');
  const sh = ss.getSheetByName('INGRESOS');

  Logger.log('🗑️ BORRANDO CORTES EQUIVOCADOS DE JUNIO 2026');
  Logger.log('');

  try {
    // Obtener los datos ANTES de borrar
    let datos = sh.getDataRange().getValues();
    Logger.log('Total de filas en Sheet: ' + datos.length);
    Logger.log('');

    // Identificar las filas a borrar por criterio (no por número de fila)
    const filasABorrar = [];
    const tz = Session.getScriptTimeZone();

    const formatearFecha = function(celda) {
      if (celda instanceof Date) {
        return Utilities.formatDate(celda, tz, 'dd-MM-yyyy');
      }
      return String(celda || '').trim();
    };

    // BUSCAR las filas que cumplan los criterios
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const fechaFila = formatearFecha(fila[1]);
      const sucursalFila = String(fila[2] || '').trim();
      const turnoFila = String(fila[3] || '').trim();

      // Criterios a buscar
      if ((fechaFila === '20-06-2026' || fechaFila === '22-06-2026' || fechaFila === '23-06-2026') &&
          sucursalFila.includes('CULTURA') &&
          (turnoFila.includes('MAÑANA') || turnoFila.includes('TARDE'))) {
        filasABorrar.push({
          num: i + 1,
          fecha: fechaFila,
          sucursal: sucursalFila,
          turno: turnoFila
        });
      }
    }

    if (filasABorrar.length === 0) {
      Logger.log('⚠️ NO se encontraron filas que cumplan los criterios.');
      Logger.log('Esto es extraño, verifica los datos manualmente.');
      return;
    }

    Logger.log('Encontradas ' + filasABorrar.length + ' filas para borrar:');
    filasABorrar.forEach(f => {
      Logger.log('  Fila ' + f.num + ': ' + f.fecha + ' | ' + f.sucursal + ' | ' + f.turno);
    });
    Logger.log('');

    // Borrar de ATRÁS hacia ADELANTE para que no se corra el índice
    filasABorrar.sort((a, b) => b.num - a.num);

    Logger.log('🗑️ Borrando filas (de atrás hacia adelante):');
    for (let f of filasABorrar) {
      Logger.log('  Eliminando fila ' + f.num + '...');
      sh.deleteRow(f.num);
      Logger.log('    ✓ Fila ' + f.num + ' eliminada');
      Utilities.sleep(500); // pequeña pausa para evitar race conditions
    }

    Logger.log('');
    Logger.log('✅ COMPLETADO: Se borraron ' + filasABorrar.length + ' filas');
    Logger.log('');

    // Verificar que se borraron
    const datosAfter = sh.getDataRange().getValues();
    Logger.log('Total de filas después: ' + datosAfter.length + ' (antes eran ' + datos.length + ')');

    if (datosAfter.length < datos.length) {
      Logger.log('✅ Las filas fueron eliminadas exitosamente.');
      Logger.log('⏳ Luna Smart debería actualizar automáticamente...');
    } else {
      Logger.log('⚠️ Parece que las filas NO se borraron. Verifica permisos.');
    }

  } catch(e) {
    Logger.log('❌ ERROR durante ejecución: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
}

// ── LIMPIAR CACHÉ DE APPS SCRIPT ────────────────────────────────────────────
function actualizarAppsScript() {
  // Limpiar todos los cachés disponibles
  const cache = CacheService.getScriptCache();
  cache.removeAll(['datos', 'ingresos', 'cache', 'facturas', 'clientes', 'proveedores']);

  // Forzar que Google Sheets escriba todos los cambios al disco
  SpreadsheetApp.flush();

  Logger.log('✅ Cache de Apps Script limpiado completamente');
  Logger.log('✅ Datos sincronizados con Google Sheets');
  Logger.log('');
  Logger.log('PRÓXIMOS PASOS:');
  Logger.log('1. Cierra Luna Smart completamente');
  Logger.log('2. En Chrome: Cmd+Shift+R (en Mac) o Ctrl+Shift+R (en Windows/Linux)');
  Logger.log('3. O abre en incógnito: Cmd+Shift+N (en Mac)');
  Logger.log('4. Verifica que desaparecieron los datos equivocados');
}
