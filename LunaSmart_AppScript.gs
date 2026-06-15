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
    getVENTASPARROT:               HOJAS.VENTAS_PARROT,
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
  // El frontend envía los datos dentro de body.datos. Desenvolvemos aquí
  // (con fallback a body por compatibilidad con llamadas directas/planas).
  const datos = body.datos || body;

  switch (accion) {
    case 'registrarIngreso':           return _registrarIngreso(datos);
    case 'registrarIngresoCompleto':   return _registrarIngresoCompleto(datos);
    case 'actualizarIngreso':          return _actualizarIngreso(datos);
    case 'borrarIngreso':              return _borrarIngreso(datos);
    case 'registrarFactura':           return _registrarFactura(datos);
    case 'registrarFacturaCompleta':   return _registrarFacturaCompleta(datos);
    case 'actualizarFactura':          return _actualizarFactura(datos);
    case 'borrarFactura':              return _borrarFactura(datos);
    case 'registrarArticuloDetalle':   return _registrarArticuloDetalle(datos);
    case 'registrarProveedor':         return _registrarProveedor(datos);
    case 'actualizarProveedor':        return _actualizarProveedor(datos);
    case 'actualizarCatalogoArticulo': return _actualizarCatalogoArticulo(datos);
    case 'registrarCliente':           return _registrarCliente(datos);
    case 'sincronizarParrot':          return _sincronizarParrot(body.sucursal || datos.sucursal || 'CASA DE LA CULTURA', body.desde || datos.desde, body.hasta || datos.hasta);
    case 'registrarCatalogoArticulo':  return _registrarCatalogoArticulo(datos);
    default: return _err('Acción desconocida: ' + accion);
  }
}

// ── REGISTRAR INGRESO + DETALLE DE VENTA ───────────────────────────────────
// Escribe el corte en INGRESOS y, si vienen, las líneas de qué se vendió en
// INGRESOS DETALLES (ligadas por ID_INGRESO en col B). Para conciliar manual.
function _registrarIngresoCompleto(b) {
  try {
    var sh = _getSheet(HOJAS.INGRESOS);
    var id = _nextId(HOJAS.INGRESOS, 'INGRESOS');
    var efectivo      = parseFloat(b.efectivo      || 0);
    var tarjeta       = parseFloat(b.tarjeta       || 0);
    var transferencia = parseFloat(b.transferencia || 0);
    var rappi         = parseFloat(b.rappi         || 0);
    var total         = efectivo + tarjeta + transferencia + rappi;
    var inicioCaja    = parseFloat(b.inicioCaja    || 0);
    var retiros       = parseFloat(b.retiros       || 0);
    var ventaTotal    = parseFloat(b.ventaTotal    || total);
    var fecha         = b.fecha || _fechaHoy();

    _escribirFila(sh, [
      id, fecha, b.sucursal || '', b.cliente || '',
      inicioCaja, retiros, parseFloat(b.depositos || 0),
      efectivo, tarjeta, transferencia, rappi,
      total, ventaTotal, total - ventaTotal,
      b.observaciones || 'MANUAL',
    ]);

    // Detalle de venta → INGRESOS DETALLES
    var nDet = 0;
    var detalles = b.detalles || [];
    if (detalles.length) {
      var shD = _getSheet(HOJAS.ING_DETALLES);
      var filas = detalles.filter(function(d){ return d && d.articulo; }).map(function(d){
        var cant = parseFloat(d.cantidad || 0) || 0;
        var precio = parseFloat(d.precio || 0) || 0;
        return [
          Utilities.getUuid().substring(0, 8), // A ID_CONCEPTO
          id,                                   // B ID_INGRESOS (liga al corte)
          fecha,                                // C FECHA
          d.articulo,                           // D ARTICULO
          cant,                                 // E CANTIDAD
          precio,                               // F PRECIO UNIT
          cant * precio,                        // G SUBTOTAL_LINEA
          '',                                   // H ¿APLICA IVA?
          0,                                    // I IVA MONTO
        ];
      });
      if (filas.length) {
        var filaInicio = _siguienteFilaLibre(shD, 4); // col D ARTICULO como indicador
        shD.getRange(filaInicio, 1, filas.length, 9).setValues(filas);
        nDet = filas.length;
      }
    }

    return _json({ status: 'ok', idIngreso: id, detalles: nDet });
  } catch (e) {
    return _err(e.message);
  }
}

// ── ACTUALIZAR / BORRAR INGRESO (corte) ────────────────────────────────────
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

    // Actualizar cols B..O (fecha..observaciones)
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

    // Reemplazar detalle de venta (INGRESOS DETALLES col B = ID_INGRESOS)
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

// Asigna ID a los cortes históricos que no tienen (para poder editarlos)
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

// ── REGISTRAR FACTURA COMPLETA (cabecera + todos los artículos en 1 escritura) ──
// Mucho más rápido que registrar artículo por artículo.
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
      shD.getRange(filaInicio, 1, filas.length, 9).setValues(filas);  // UNA sola escritura
      // Costos dinámicos en lote (lee el catálogo 1 vez)
      var mapa = {};
      lineas.forEach(function(l){ if (l.articulo) mapa[String(l.articulo).trim().toLowerCase()] = parseFloat(l.precioUnit) || 0; });
      _actualizarCostosDinamicos(mapa);
    }
    return _json({ status: 'ok', idFactura: id });
  } catch(e) {
    return _err(e.message);
  }
}

// Actualiza el costo dinámico de varios artículos leyendo el catálogo una sola vez
function _actualizarCostosDinamicos(mapaArticuloPrecio) {
  try {
    var sh = _getSheet(HOJAS.CATALOGO);
    var datos = sh.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      var nombre = String(datos[i][0]).trim().toLowerCase();
      if (mapaArticuloPrecio.hasOwnProperty(nombre) && mapaArticuloPrecio[nombre] > 0) {
        sh.getRange(i + 1, 3).setValue(mapaArticuloPrecio[nombre]); // col C = Costo Dinámico
      }
    }
  } catch(_) {}
}

// ── ACTUALIZAR FACTURA (editar cabecera + reemplazar artículos) ────────────
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

    // Actualizar cabecera (cols B..G)
    sh.getRange(fila, 2, 1, 6).setValues([[
      b.fecha     || vals[fila-1][1],
      b.unidad    || '',
      b.proveedor || '',
      b.folio     || '',
      vals[fila-1][5] || '',                 // FOTO (se conserva)
      parseFloat(b.total || 0)
    ]]);

    // Reemplazar artículos: borrar los existentes de esta factura y reinsertar
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

// ── BORRAR FACTURA (cabecera + sus artículos) ──────────────────────────────
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
// Actualizar proveedor existente (por ID en col A)
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

// Actualizar artículo del Catálogo Maestro (por nombre clave)
function _actualizarCatalogoArticulo(b) {
  try {
    var clave = String(b.articuloKey || b.articulo || '').trim().toLowerCase();
    if (!clave) return _err('Falta el artículo');
    var sh = _getSheet(HOJAS.CATALOGO);
    var vals = sh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toLowerCase() === clave) {
        var costo = parseFloat(b.costoBase || vals[i][1] || 0);
        var merma = parseFloat(b.merma || 0);
        var costoFinal = merma > 0 ? costo / (1 - merma / 100) : costo;
        // A=articulo B=costoBase C=costoDinamico(conserva) D=cantidad(conserva)
        sh.getRange(i + 1, 1, 1, 2).setValues([[b.articulo || vals[i][0], costo]]);
        sh.getRange(i + 1, 5, 1, 6).setValues([[
          b.unidad || '', merma, costoFinal,
          b.categoria || '', b.subcategoria || '', b.proveedor || '',
        ]]);
        return _json({ status: 'ok' });
      }
    }
    return _err('Artículo no encontrado: ' + clave);
  } catch (e) { return _err(e.message); }
}

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
    // Usar prefijo SL- para continuar la numeración histórica de proveedores
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

// ── DIAGNÓSTICO PARROT ─────────────────────────────────────────────────────
// Ejecuta esta función desde el editor y copia TODO el "Registro de ejecución".
// Sirve para ver qué devuelve la API de Parrot y si trae el detalle de productos
// vendidos (para poder construir "Ventas por artículo / Artículos estrella").
function diagnosticarParrot() {
  // Ventana de 24h (la API exige máximo 48h entre start y end)
  var hoy = new Date();
  var ayer = new Date(hoy - 24 * 60 * 60 * 1000);
  var tz = 'America/Mexico_City';
  var ini = Utilities.formatDate(ayer, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
  var fin = Utilities.formatDate(hoy,  tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
  var qParams = '?startTimestamp=' + encodeURIComponent(ini) + '&endTimestamp=' + encodeURIComponent(fin) +
                '&storeUUID=' + PARROT_STORE_UUID + '&pageSize=5';

  // Endpoints oficiales (base: api.parrot.rest/external)
  var endpoints = [
    { nom: 'List Orders',           url: PARROT_BASE_URL + '/v1/orders' + qParams },
    { nom: 'List Order Items v2',   url: PARROT_BASE_URL + '/v2/order-items' + qParams },
    { nom: 'List Order Items v1',   url: PARROT_BASE_URL + '/v1/order-items' + qParams },
    { nom: 'List Cashier Sessions', url: PARROT_BASE_URL + '/v1/cashier-sessions' + qParams },
  ];
  var headers = { 'Authorization': 'Bearer ' + PARROT_API_KEY };

  endpoints.forEach(function(ep) {
    Logger.log('═══════════════════════════════════════════');
    Logger.log('PROBANDO: ' + ep.nom);
    Logger.log(ep.url);
    try {
      var resp = UrlFetchApp.fetch(ep.url, { headers: headers, muteHttpExceptions: true });
      var code = resp.getResponseCode();
      Logger.log('HTTP ' + code);
      var body = resp.getContentText();
      if (code !== 200) {
        Logger.log('Respuesta: ' + body.substring(0, 300));
        return;
      }
      var data = JSON.parse(body);
      var items = data.data || data.results || (Array.isArray(data) ? data : []);
      Logger.log('✅ Registros: ' + items.length);
      if (items.length > 0) {
        Logger.log('── Campos: ' + Object.keys(items[0]).join(', '));
        Logger.log('── PRIMER REGISTRO COMPLETO:');
        Logger.log(JSON.stringify(items[0], null, 2).substring(0, 2500));
      }
    } catch(e) {
      Logger.log('Error: ' + e.message);
    }
  });
  Logger.log('═══════════════════════════════════════════');
  Logger.log('FIN. Cópiame TODO este registro.');
}

// ── PARROT POS — API api.parrot.rest/external ──────────────────────────────
// Sucursales con Parrot: CASA DE LA CULTURA (principal), HELFY FÜ (multimarca).
// Modelo (igual que AppSheet):
//   cashier-session  → 1 fila en INGRESOS (corte: pagos + venta + diferencia)
//   order-items      → filas en INGRESOS DETALLES con el ID_INGRESO del corte

// GET autenticado a Parrot. Ventana [ini, fin] máx 48h. pageSize máx 100.
function _parrotGet(path, ini, fin, pageSize, page) {
  var tz = 'America/Mexico_City';
  var qs = '?startTimestamp=' + encodeURIComponent(Utilities.formatDate(ini, tz, "yyyy-MM-dd'T'HH:mm:ssXXX"))
         + '&endTimestamp='   + encodeURIComponent(Utilities.formatDate(fin, tz, "yyyy-MM-dd'T'HH:mm:ssXXX"))
         + '&storeUUID=' + PARROT_STORE_UUID
         + '&pageSize=' + Math.min(pageSize || 100, 100)
         + '&page=' + (page || 0);
  var resp = UrlFetchApp.fetch(PARROT_BASE_URL + path + qs, {
    headers: { 'Authorization': 'Bearer ' + PARROT_API_KEY },
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Parrot ' + path + ' HTTP ' + code + ': ' + resp.getContentText().substring(0, 150));
  }
  var d = JSON.parse(resp.getContentText());
  return d.data || d.results || [];
}

function _isoToDate(iso, finDelDia) {
  var p = String(iso).split('-');
  return new Date(parseInt(p[0],10), parseInt(p[1],10) - 1, parseInt(p[2],10),
                  finDelDia ? 23 : 0, finDelDia ? 59 : 0, 0);
}

// Sincroniza Parrot al modelo INGRESOS + INGRESOS DETALLES. Chunks de 24h.
function _sincronizarParrot(sucursal, desdeISO, hastaISO) {
  sucursal = sucursal || 'CASA DE LA CULTURA';
  try {
    var fin = hastaISO ? _isoToDate(hastaISO, true)  : new Date();
    var ini = desdeISO ? _isoToDate(desdeISO, false) : new Date(fin - 2 * 24 * 60 * 60 * 1000);

    var shIng = _getSheet(HOJAS.INGRESOS);
    var shDet = _getSheet(HOJAS.ING_DETALLES);
    var tz = Session.getScriptTimeZone();

    // Turno por horario MX: 8:00–13:59 = MAÑANA, 14:00+ = TARDE
    var turnoDe = function(d){
      var h = parseInt(Utilities.formatDate(d, 'America/Mexico_City', 'HH'), 10);
      return h < 14 ? 'TURNO MAÑANA' : 'TURNO TARDE';
    };

    // Dedup: sesiones ya importadas (UUID en OBSERVACIONES como "PARROT:uuid")
    var sesVistas = {};
    shIng.getDataRange().getValues().slice(1).forEach(function(r){
      var m = String(r[14] || '').match(/PARROT:([a-f0-9-]+)/i);
      if (m) sesVistas[m[1]] = r[0];   // uuid → ID_INGRESO
    });
    // Dedup: items ya importados (UUID en ID_CONCEPTO = col A)
    var itemVistos = {};
    shDet.getDataRange().getValues().slice(1).forEach(function(r){ if (r[0]) itemVistos[String(r[0])] = true; });

    var nCortes = 0, nItems = 0;
    var cursor = new Date(ini);
    while (cursor < fin) {
      var chunkFin = new Date(Math.min(cursor.getTime() + 24*60*60*1000, fin.getTime()));

      // Negocio cerrado los DOMINGOS → saltar (no hay ventas, evita llamadas)
      // 'u' = día ISO de la semana (1=Lun ... 7=Dom) en hora de México
      if (Utilities.formatDate(cursor, 'America/Mexico_City', 'u') === '7') {
        cursor = chunkFin;
        continue;
      }

      // 1) CORTES (cashier-sessions) → INGRESOS. Mapa turno → ID_INGRESO del día.
      var sesiones = [];
      try { sesiones = _parrotGet('/v1/cashier-sessions', cursor, chunkFin, 50, 0); }
      catch(e){ Logger.log('sesiones: ' + e.message); }
      Utilities.sleep(4500);

      var corteDeTurno = {};   // 'TURNO MAÑANA' → ID_INGRESO, 'TURNO TARDE' → ID_INGRESO
      sesiones.forEach(function(s){
        var inicioSes = new Date(s.startedAt || s.finishedAt);
        var turno = turnoDe(inicioSes);
        if (s.uuid && sesVistas[s.uuid]) { if (!corteDeTurno[turno]) corteDeTurno[turno] = sesVistas[s.uuid]; return; }
        var pay = {};
        (s.sessionByPaymentType || []).forEach(function(p){ pay[p.paymentType] = (pay[p.paymentType]||0) + (p.reportedAmount||0); });
        var cm = s.cashMovements || {};
        var efectivo = pay['CASH'] || 0;
        var tarjeta  = (pay['DEBIT_CARD']||0) + (pay['CREDIT_CARD']||0) + (pay['PAY']||0);
        var transfer = pay['THIRD_PARTY'] || 0;
        var ventaParrot = (s.sales && s.sales.totalSales) || cm.expectedAmount || 0;
        var declarado = cm.reportedAmount || 0;
        var dif = (cm.differenceAmount != null) ? cm.differenceAmount : (declarado - ventaParrot);
        var estado = Math.abs(dif) < 1 ? 'CUADRA' : (dif < 0 ? 'FALTANTE $' + Math.abs(dif) : 'SOBRANTE $' + dif);
        var idIng = Utilities.getUuid().substring(0, 8);
        // ISO yyyy-MM-dd para que Google Sheets NO invierta día/mes
        var fecha = Utilities.formatDate(new Date(s.finishedAt || s.startedAt), tz, 'yyyy-MM-dd');
        _escribirFila(shIng, [
          idIng, fecha, sucursal, turno,
          cm.startingAmount || 0, cm.withdrawals || 0, cm.deposits || 0,
          efectivo, tarjeta, transfer, 0,
          declarado, ventaParrot, dif,
          estado + ' | PARROT:' + s.uuid
        ]);
        sesVistas[s.uuid] = idIng;
        if (!corteDeTurno[turno]) corteDeTurno[turno] = idIng;
        nCortes++;
      });

      // Fallback: si solo hay un corte en el día, sirve para cualquier turno
      var unicoCorte = null;
      for (var k in corteDeTurno) { unicoCorte = unicoCorte || corteDeTurno[k]; }

      // 2) VENTAS POR ARTÍCULO → INGRESOS DETALLES (ligado al corte del MISMO turno)
      if (unicoCorte) {
        var page = 0, hayMas = true, guard = 0;
        while (hayMas && guard < 25) {
          var items = [];
          try { items = _parrotGet('/v2/order-items', cursor, chunkFin, 100, page); }
          catch(e){ Logger.log('items p' + page + ': ' + e.message); break; }
          Utilities.sleep(4500);
          items.forEach(function(it){
            if (it.uuid && itemVistos[it.uuid]) return;  // ya importado
            var t = new Date(it.createdAt);
            var idIng = corteDeTurno[turnoDe(t)] || unicoCorte;
            var fecha = Utilities.formatDate(t, tz, 'yyyy-MM-dd');  // ISO (evita inversión)
            var total = parseFloat(it.total) || 0;
            _escribirFila(shDet, [
              it.uuid || Utilities.getUuid().substring(0,8),  // A=ID_CONCEPTO (dedup)
              idIng,                                           // B=ID_INGRESOS (corte del turno)
              fecha,                                           // C=FECHA
              it.itemName || '',                               // D=ARTICULO
              parseFloat(it.quantity) || 0,                    // E=CANTIDAD
              parseFloat(it.unitPrice) || 0,                   // F=PRECIO UNIT
              total,                                           // G=SUBTOTAL
              'NO',                                            // H=APLICA IVA
              parseFloat(it.taxAmount) || 0,                   // I=IVA
              total                                            // J=TOTAL
            ]);
            if (it.uuid) itemVistos[it.uuid] = true;
            nItems++;
          });
          hayMas = items.length === 100;
          page++;
        }
      }
      cursor = chunkFin;
    }

    return _json({ status: 'ok', msg: 'Parrot sincronizado',
                   registros: nCortes + nItems, cortes: nCortes, articulos: nItems });
  } catch (e) {
    return _err('Parrot: ' + e.message);
  }
}

// ════════════════════════════════════════════════════════════════
// FACTURAS rechaza setNumberFormat por un formato pegado en la columna.
// Esta función: (1) pone el spreadsheet en es_MX para que las fechas NUEVAS
// salgan dd/mm, (2) limpia el formato de la columna y reescribe las fechas
// reales (serial) como texto ISO yyyy-mm-dd. Ejecútala UNA VEZ.
// ════════════════════════════════════════════════════════════════
function arreglarFechasFacturas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var tz = ss.getSpreadsheetTimeZone();
  var log = [];

  // 1) Locale a México (las fechas NUEVAS saldrán dd/mm en vez de m/d)
  try { ss.setSpreadsheetLocale('es_MX'); log.push('locale → es_MX ✅'); }
  catch (e) { log.push('locale ERROR: ' + e.message); }

  var sh = ss.getSheetByName('FACTURAS');
  var n = sh.getLastRow() - 1;
  var rng = sh.getRange(2, 2, n, 1);
  var vals = rng.getValues();   // fechas reales como Date objects

  var out = vals.map(function(r){
    var v = r[0];
    if (v === '' || v == null) return [''];
    var d = (Object.prototype.toString.call(v) === '[object Date]') ? v : new Date(v);
    if (isNaN(d.getTime())) return [v];
    return [Utilities.formatDate(d, tz, 'yyyy-MM-dd')];
  });

  // 2) Limpiar el formato pegado y reescribir como texto ISO
  var ok = false;
  try {
    rng.clearFormat();
    SpreadsheetApp.flush();
    rng.setNumberFormat('@');
    rng.setValues(out);
    ok = true;
    log.push('FACTURAS → ' + n + ' fechas a ISO texto ✅');
  } catch (e) {
    log.push('estrategia texto falló: ' + e.message);
  }

  // 3) Fallback: dejarlas como fechas reales con formato dd/mm/yyyy
  if (!ok) {
    try {
      rng.clearFormat();
      SpreadsheetApp.flush();
      rng.setNumberFormat('dd/mm/yyyy');
      log.push('FACTURAS → formato dd/mm/yyyy (fallback) ✅');
    } catch (e2) {
      log.push('fallback dd/mm falló: ' + e2.message);
    }
  }

  Logger.log('Resultado arreglarFechasFacturas:\n  ' + log.join('\n  '));
}

// ════════════════════════════════════════════════════════════════
// FECHAS: normaliza el FORMATO de las columnas de fecha a ISO yyyy-mm-dd
// Soluciona que Google las devuelva en formato gringo M/D/YYYY (que el
// sistema malinterpretaba: 6/12 = 12-jun se leía como 6-dic).
// Solo cambia el FORMATO de despliegue, NO reescribe los valores → 100% seguro.
// Ejecuta esta función UNA VEZ desde el editor.
// ════════════════════════════════════════════════════════════════
function formatearFechasISO() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var tz = ss.getSpreadsheetTimeZone();
  // [nombreHoja, columna 1-based de la FECHA]
  var cols = [
    ['INGRESOS', 2],
    ['FACTURAS', 2],
    ['ARTICULOS DETALLES', 3],
    ['INGRESOS DETALLES', 3],
    ['CONCILIACION', 1]
  ];
  var log = [];
  cols.forEach(function(c){
    try {
      var sh = ss.getSheetByName(c[0]);
      if (!sh) { log.push(c[0] + ': (no existe)'); return; }
      var n = sh.getLastRow() - 1;
      if (n < 1) { log.push(c[0] + ': vacía'); return; }
      var rng = sh.getRange(2, c[1], n, 1);
      var vals = rng.getValues();
      var conv = 0;
      var out = vals.map(function(r){
        var v = r[0];
        if (v === '' || v == null) return [''];
        // El serial (Date real) es la verdad — lo convertimos a ISO sin ambigüedad
        var d = (Object.prototype.toString.call(v) === '[object Date]') ? v : null;
        if (!d) {
          // por si alguna quedó como texto: intentar parsear
          var s = String(v).trim();
          if (/^\d{4}-\d{2}-\d{2}/.test(s)) return [s.substring(0,10)]; // ya ISO
          return [v]; // se queda igual, no la tocamos
        }
        conv++;
        return [Utilities.formatDate(d, tz, 'yyyy-MM-dd')];
      });
      // Formato texto a TODA la columna (incluye filas futuras) para que
      // las facturas/ingresos nuevos también se guarden como ISO sin reinterpretar.
      sh.getRange(2, c[1], sh.getMaxRows() - 1, 1).setNumberFormat('@');
      rng.setValues(out);
      log.push(c[0] + ': ' + conv + '/' + n + ' fechas → ISO texto ✅');
    } catch (e) {
      log.push(c[0] + ': ERROR ' + e.message);
    }
  });
  Logger.log('Resultado formatearFechasISO:\n  ' + log.join('\n  '));
}

// Backfill manual desde el editor. Ej: sincronizarParrotDias(2)
function sincronizarParrotDias(dias) {
  dias = dias || 2;
  var hasta = new Date();
  var desde = new Date(hasta - dias * 24 * 60 * 60 * 1000);
  var iso = function(d){ return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); };
  var r = _sincronizarParrot('CASA DE LA CULTURA', iso(desde), iso(hasta));
  Logger.log(r.getContent());
}

// BACKFILL: jala un rango de fechas completo. EDITA DESDE/HASTA y ejecuta.
// Máximo ~3 semanas por corrida (límite de 6 min de Apps Script).
// ════════════════════════════════════════════════════════════════
// ESTANDARIZAR SUCURSALES — deja UN solo nombre por sucursal y pone
// menús desplegables en INGRESOS.SUCURSAL y FACTURAS.UNIDAD.
// - Cortes de Parrot (OBS contiene PARROT) → CASA DE LA CULTURA
// - Facturas con unidad vacía → CASA DE LA CULTURA
// - Estandariza variantes/typos a los 5 nombres oficiales
// Ejecútala UNA VEZ.
// ════════════════════════════════════════════════════════════════
var SUCURSALES_OFICIALES = ['COFFEE & ROASTERS','CASA DE LA CULTURA','HELFY FÜ','BARBACOA Y MENUDO','EVENTOS'];

function _canonSucursal(v, vaciaDefault) {
  var t = String(v == null ? '' : v).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
  if (!t) return vaciaDefault || '';
  if (/CULTURA/.test(t)) return 'CASA DE LA CULTURA';
  if (/HELFY/.test(t)) return 'HELFY FÜ';
  if (/BARBACOA|MENUDO|BENJAMIN|JAIME/.test(t)) return 'BARBACOA Y MENUDO';
  if (/EVENTO/.test(t)) return 'EVENTOS';
  if (/COFFEE|COOFFEE|ROASTER/.test(t)) return 'COFFEE & ROASTERS';
  if (/SUENO DE LUNA/.test(t)) return 'COFFEE & ROASTERS';
  return v; // desconocido: dejar igual
}

// Renombra el encabezado de FACTURAS: "UNIDAD DE NEGOCIO" → "SUCURSAL"
function renombrarEncabezadoSucursal() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('FACTURAS');
  try {
    sh.getRange('C1').setValue('SUCURSAL');
    Logger.log('✅ FACTURAS!C1 renombrado a "SUCURSAL"');
  } catch (e) {
    Logger.log('❌ No se pudo por script: ' + e.message +
               '\n→ Hazlo a mano: doble clic en la celda C1 de FACTURAS y escribe SUCURSAL.');
  }
}

function estandarizarSucursales() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var log = [];
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(SUCURSALES_OFICIALES, true).setAllowInvalid(true).build();

  // ── INGRESOS: cortes de Parrot → CASA DE LA CULTURA (col C=3, OBS col O=15) ──
  try {
    var shI = ss.getSheetByName('INGRESOS');
    var nI = shI.getLastRow() - 1;
    if (nI > 0) {
      var rngS = shI.getRange(2, 3, nI, 1);
      var suc = rngS.getValues();
      var obs = shI.getRange(2, 15, nI, 1).getValues();
      var ch = 0;
      for (var i = 0; i < suc.length; i++) {
        var o = String(obs[i][0] || '').toUpperCase();
        if (/PARROT/.test(o) && suc[i][0] !== 'CASA DE LA CULTURA') { suc[i][0] = 'CASA DE LA CULTURA'; ch++; }
      }
      try { rngS.clearDataValidations(); } catch (e) {}
      rngS.setValues(suc);
      log.push('INGRESOS: ' + ch + ' cortes Parrot → CASA DE LA CULTURA ✅');
      try { rngS.setDataValidation(rule); log.push('INGRESOS: dropdown ✅'); }
      catch (e) { log.push('INGRESOS dropdown omitido (' + e.message + ')'); }
    }
  } catch (e) { log.push('INGRESOS ERROR: ' + e.message); }

  // ── FACTURAS: estandarizar UNIDAD (col C=3), vacías → CASA DE LA CULTURA ──
  try {
    var shF = ss.getSheetByName('FACTURAS');
    var nF = shF.getLastRow() - 1;
    if (nF > 0) {
      var rngU = shF.getRange(2, 3, nF, 1);
      var uni = rngU.getValues();
      var ch2 = 0;
      for (var j = 0; j < uni.length; j++) {
        var canon = _canonSucursal(uni[j][0], 'CASA DE LA CULTURA');
        if (uni[j][0] !== canon) { uni[j][0] = canon; ch2++; }
      }
      try { rngU.clearDataValidations(); } catch (e) {}
      rngU.setValues(uni);
      log.push('FACTURAS: ' + ch2 + ' unidades estandarizadas ✅');
      try { rngU.setDataValidation(rule); log.push('FACTURAS: dropdown ✅'); }
      catch (e) { log.push('FACTURAS dropdown omitido (' + e.message + ')'); }
    }
  } catch (e) { log.push('FACTURAS ERROR: ' + e.message); }

  Logger.log('Resultado estandarizarSucursales:\n  ' + log.join('\n  '));
}

function backfillParrot() {
  var DESDE = '2026-06-01';   // ← edita: primer día a sincronizar
  // HASTA = hoy automáticamente (cubre todos los días hasta la fecha actual)
  var HASTA = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var r = _sincronizarParrot('CASA DE LA CULTURA', DESDE, HASTA);
  Logger.log(r.getContent());
}

// LIMPIEZA: borra TODOS los cortes importados de Parrot (INGRESOS con "PARROT:")
// y sus productos en INGRESOS DETALLES. Útil para re-sincronizar limpio.
// Borra SOLO los cortes fantasma del primer sync (formato viejo "PARROT-252",
// mal fechados en ago/sep/oct). NO toca los cortes buenos ("PARROT:uuid").
function borrarFantasmasParrot() {
  var sh = _getSheet(HOJAS.INGRESOS);
  var shD = _getSheet(HOJAS.ING_DETALLES);
  var vals = sh.getDataRange().getValues();
  var ids = {}, filas = [];
  for (var i = 1; i < vals.length; i++) {
    if (/PARROT-\d/i.test(String(vals[i][14] || ''))) {
      var id = String(vals[i][0]).trim();
      if (id) ids[id] = true;
      filas.push(i + 1);
    }
  }
  // borrar detalles ligados
  var dv = shD.getDataRange().getValues(), fd = [];
  for (var j = 1; j < dv.length; j++) { var r = String(dv[j][1]).trim(); if (r && ids[r]) fd.push(j + 1); }
  fd.sort(function(a,b){return b-a;}).forEach(function(f){ shD.deleteRow(f); });
  filas.sort(function(a,b){return b-a;}).forEach(function(f){ sh.deleteRow(f); });
  Logger.log('🗑️ Borrados ' + filas.length + ' cortes fantasma (PARROT-### viejo) y ' + fd.length + ' detalles.');
}

function borrarCortesParrot() {
  var shIng = _getSheet(HOJAS.INGRESOS);
  var shDet = _getSheet(HOJAS.ING_DETALLES);

  // 1) Identificar cortes de Parrot y sus ID_INGRESO
  var ingVals = shIng.getDataRange().getValues();
  var idsParrot = {};
  var filasIng = [];
  for (var i = 1; i < ingVals.length; i++) {
    // Atrapa ambos formatos: "PARROT:uuid" (nuevo) y "PARROT-252" (viejo)
    if (/PARROT[:\-]/i.test(String(ingVals[i][14] || ''))) {
      var idC = String(ingVals[i][0]).trim();
      if (idC) idsParrot[idC] = true;   // solo IDs no vacíos (evita borrar detalles ajenos)
      filasIng.push(i + 1);             // la fila del corte se borra igual
    }
  }
  // 2) Borrar detalles ligados a esos ID_INGRESO (col B), nunca con ID vacío
  var detVals = shDet.getDataRange().getValues();
  var filasDet = [];
  for (var j = 1; j < detVals.length; j++) {
    var ref = String(detVals[j][1]).trim();
    if (ref && idsParrot[ref]) filasDet.push(j + 1);
  }
  // Borrar de abajo hacia arriba (para no correr índices)
  filasDet.sort(function(a,b){return b-a;}).forEach(function(f){ shDet.deleteRow(f); });
  filasIng.sort(function(a,b){return b-a;}).forEach(function(f){ shIng.deleteRow(f); });

  Logger.log('🗑️ Borrados ' + filasIng.length + ' cortes y ' + filasDet.length + ' productos de Parrot.');
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

// ══════════════════════════════════════════════════════════════════════════
// CATEGORIZACIÓN Y LISTAS DESPLEGABLES DEL CATÁLOGO MAESTRO
// ══════════════════════════════════════════════════════════════════════════

// ── 1) ASIGNAR CATEGORÍA/SUBCATEGORÍA A ARTÍCULOS SIN CLASIFICAR ───────────
// Ejecuta esta función UNA VEZ desde el editor. Asigna las categorías a los
// artículos que estaban vacíos (según clasificación revisada).
function asignarCategoriasFaltantes() {
  var sh = _getSheet(HOJAS.CATALOGO);
  var datos = sh.getDataRange().getValues();

  // Mapa ARTICULO → [CATEGORIA, SUBCATEGORIA]
  var MAP = {
    'ENDULZANTE NATURAL STEVIA':      ['COSTO DE VENTA', 'ABARROTES'],
    'GARRAFÓN DE AGUA 19LTS':         ['COSTO DE VENTA', 'AGUAS Y REFRESCOS'],
    'MANCHEGO DURANGUEÑO':            ['COSTO DE VENTA', 'LACTEOS'],
    'TAPATÍO BOWL.':                  ['COSTO DE VENTA', 'ADEREZOS'],
    'CARM SYRP':                      ['COSTO DE VENTA', 'CAFÉ E INSUMOS'],
    'PIZZAS':                         ['COSTO DE VENTA', 'CONGELADOS'],
    'ESPONJA CH':                     ['GASTO OPERATIVO', 'LIMPIEZA'],
    'SALSA HUTS':                     ['COSTO DE VENTA', 'ADEREZOS'],
    'CIABATTA RÚSTICA':               ['COSTO DE VENTA', 'PANADERIA Y REPOSTERIA'],
    'ALITAS':                         ['COSTO DE VENTA', 'PROTEINAS'],
    'PECHUGA C HUESO':                ['COSTO DE VENTA', 'PROTEINAS'],
    'CIABATTA MULT':                  ['COSTO DE VENTA', 'PANADERIA Y REPOSTERIA'],
    'QUESO AMERICANO':                ['COSTO DE VENTA', 'LACTEOS'],
    'MANZANA GRANNY':                 ['COSTO DE VENTA', 'FRUTAS Y VERDURAS'],
    'NUEX CORAZÓN':                   ['COSTO DE VENTA', 'ABARROTES'],
    'HORNEADO ESP':                   ['COSTO DE VENTA', 'PANADERIA Y REPOSTERIA'],
    'BAGUEL':                         ['COSTO DE VENTA', 'PANADERIA Y REPOSTERIA'],
    'QUESO DURANG':                   ['COSTO DE VENTA', 'LACTEOS'],
    'MEDIA BAGUETTE':                 ['COSTO DE VENTA', 'PANADERIA Y REPOSTERIA'],
    'SERVICIO A DOMICILIO':           ['GASTO OPERATIVO', 'REPARTIDORES INDEPENDIENTES'],
    // Estos 3 quedan a tu criterio (los dejo señalados en el log, NO se tocan):
    // 'PIÑATAS CON DULCE PARA 24 PERSONAS', 'PREMIO PARA CONCURSO', 'Regular'
  };

  var asignados = 0, pendientes = [];
  for (var i = 1; i < datos.length; i++) {
    var art = String(datos[i][0]).trim();
    if (!art) continue;
    var catActual = String(datos[i][7] || '').trim();
    if (catActual) continue; // ya tiene categoría
    if (MAP[art]) {
      sh.getRange(i + 1, 8).setValue(MAP[art][0]);  // col H = CATEGORIA
      sh.getRange(i + 1, 9).setValue(MAP[art][1]);  // col I = SUBCATEGORIA
      asignados++;
    } else {
      pendientes.push(art);
    }
  }
  Logger.log('✅ Asignados automáticamente: ' + asignados);
  Logger.log('⚠️ Quedan para revisar manualmente (' + pendientes.length + '): ' + pendientes.join(' | '));
}

// ── 2) PONER LISTAS DESPLEGABLES (validación) EN CATÁLOGO MAESTRO ──────────
// Ejecuta una vez. Pone validación de lista en la columna CATEGORIA (H).
// La SUBCATEGORIA (I) se llena en cascada con el trigger onEdit de abajo.
function configurarValidacionCatalogo() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = _getSheet(HOJAS.CATALOGO);
  var cats = _categoriasUnicas();
  var ultima = Math.max(sh.getLastRow(), 2);

  var regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(cats, true)
    .setAllowInvalid(false)
    .setHelpText('Elige una categoría de la lista')
    .build();
  // Aplicar a CATEGORIA (col 8) desde fila 2 hasta el final
  sh.getRange(2, 8, ultima - 1, 1).setDataValidation(regla);
  Logger.log('✅ Validación de CATEGORIA aplicada. Categorías: ' + cats.join(', '));
  Logger.log('ℹ️ La SUBCATEGORIA se filtra sola al elegir categoría (trigger onEdit).');
}

// Devuelve las categorías únicas desde DATOS_CATEGORIASSUBCATEGORIAS
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

// Devuelve las subcategorías de una categoría dada
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

// ── 3) CASCADA: al cambiar CATEGORIA, filtrar SUBCATEGORIA ─────────────────
// Trigger simple onEdit — se ejecuta solo al editar el Sheet.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== HOJAS.CATALOGO) return;
    var col = e.range.getColumn();
    var row = e.range.getRow();
    if (col !== 8 || row < 2) return; // solo columna CATEGORIA (H)

    var categoria = String(e.range.getValue() || '').trim();
    var celdaSub = sh.getRange(row, 9); // col I = SUBCATEGORIA
    if (!categoria) { celdaSub.clearDataValidations(); return; }

    var subs = _subcategoriasDe(categoria);
    if (subs.length > 0) {
      var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(subs, true)
        .setAllowInvalid(false)
        .build();
      celdaSub.setDataValidation(regla);
      celdaSub.clearContent(); // limpiar subcategoría anterior (era de otra categoría)
    } else {
      celdaSub.clearDataValidations();
    }
  } catch(err) {
    // Silencioso para no interrumpir la edición
  }
}
