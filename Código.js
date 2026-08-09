/*******************************************************
 * SISTEMA DE ADMINISTRACIÓN DE CONDOMINIOS GOS
 * Backend - Código.gs
 *******************************************************/

const SHEET_ID = '12TnESeCpcRM6qiJ-T_K6nr17KIqLRRCHqyxzt_5HJ9o';

const HOJAS = {
  PROPIETARIOS: ['CASA','PROPIETARIOS','IDENTIFICACION','CONTACTO 1','CONTACTO 2','CONTACTO 3','WHASTAPP','CORREO PRINCIPAL','CORREO SECUNDARIO','ALICUOTA','CEDULA','SOLVENCIA','PDF','CALLE','ESTADO','CONTRASEÑA','ROL'],
  CARGOS: ['N° REG','CASA','REFERENCIA','MONTO ($)'],
  PAGOS: ['N° REG','FECHA','CASA','DESCRIPCION','MONTO (BS)','MONTO ($)','Ref. US $'],
  ESTADO: ['CASA','PROPIETARIO(S)','RECIBOS VENCIDOS','SALDO','ESTATUS']
};

const MESES = {
  'ENERO':1,'FEBRERO':2,'MARZO':3,'ABRIL':4,'MAYO':5,'JUNIO':6,
  'JULIO':7,'AGOSTO':8,'SEPTIEMBRE':9,'OCTUBRE':10,'NOVIEMBRE':11,'DICIEMBRE':12
};

const LOGO_URL = 'https://www.appsheet.com:443/fsimage.png?appid=a1e80fbe-c271-4911-826d-05b21b90e880&datasource=google&filename=DocId%3D1qEe645o4akG7FD-H90Z8xT479Lz65SdH&signature=e17b9a139af8848ef35f37882d40f1a8a098f22c5627d8879a8c1dd967590f2f&tableprovider=google&userid=5890721';

/*******************************************************
 * PUNTO DE ENTRADA WEB
 *******************************************************/
function doGet(e) {
  const estadoSistema = obtenerEstadoSistema();
  if (estadoSistema.estado === 'CERRADO' || estadoSistema.estado === 'MANTENIMIENTO') {
    return construirPaginaBloqueo(estadoSistema);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistema de Administración de Condominios GOS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/*******************************************************
 * SETUP: Crea hojas y encabezados si no existen
 * (Ejecutar manualmente una sola vez desde el editor)
 *******************************************************/
function crearHojasYEncabezados() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Object.keys(HOJAS).forEach(function (nombre) {
    let hoja = ss.getSheetByName(nombre);
    if (!hoja) {
      hoja = ss.insertSheet(nombre);
    }
    const headers = HOJAS[nombre];
    const rango = hoja.getRange(1, 1, 1, headers.length);
    const actuales = rango.getValues()[0];
    const vacio = actuales.every(function (v) { return v === '' || v === null; });
    if (vacio) {
      rango.setValues([headers]);
      rango.setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
      hoja.setFrozenRows(1);
      try { hoja.autoResizeColumns(1, headers.length); } catch (err) {}
    }
  });
  crearHojaEstadoSistema();
  return 'Hojas y encabezados verificados/creados correctamente.';
}

/*******************************************************
 * SETUP: Crea la hoja de control de estado del sistema
 * (Se ejecuta también automáticamente dentro de crearHojasYEncabezados)
 *******************************************************/
function crearHojaEstadoSistema() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let hoja = ss.getSheetByName('ESTADO_SISTEMA');
  if (!hoja) {
    hoja = ss.insertSheet('ESTADO_SISTEMA');
  }
  const headers = ['ESTADO', 'MENSAJE'];
  const rango = hoja.getRange(1, 1, 1, headers.length);
  const actuales = rango.getValues()[0];
  const vacio = actuales.every(function (v) { return v === '' || v === null; });
  if (vacio) {
    rango.setValues([headers]);
    rango.setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) {
    hoja.getRange(2, 1, 1, 2).setValues([['ABIERTO', '']]);
  }
  return 'Hoja ESTADO_SISTEMA verificada/creada correctamente.';
}

/*******************************************************
 * Lee el estado actual del sistema (ABIERTO / CERRADO / MANTENIMIENTO)
 *******************************************************/
function obtenerEstadoSistema() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const hoja = ss.getSheetByName('ESTADO_SISTEMA');
    if (!hoja) return { estado: 'ABIERTO', mensaje: '' };
    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return { estado: 'ABIERTO', mensaje: '' };
    const idx = _indexarEncabezados(datos[0]);
    const estado = String(_val(datos[1], idx, 'ESTADO') || 'ABIERTO').trim().toUpperCase();
    const mensaje = String(_val(datos[1], idx, 'MENSAJE') || '');
    return { estado: estado, mensaje: mensaje };
  } catch (err) {
    return { estado: 'ABIERTO', mensaje: '' };
  }
}

/*******************************************************
 * Página que se muestra cuando el sistema está
 * CERRADO o en MANTENIMIENTO
 *******************************************************/
function construirPaginaBloqueo(estadoSistema) {
  const esMantenimiento = estadoSistema.estado === 'MANTENIMIENTO';
  const titulo = esMantenimiento ? 'En mantenimiento' : 'Acceso cerrado temporalmente';
  const icono = esMantenimiento ? '🛠️' : '🔒';
  const mensajeDefault = esMantenimiento
    ? 'Estamos realizando mejoras al sistema. Vuelve a intentarlo en unos minutos.'
    : 'El portal se encuentra cerrado temporalmente. Contacta a la administración para más información.';
  const mensajeFinal = (estadoSistema.mensaje && estadoSistema.mensaje.trim()) ? estadoSistema.mensaje : mensajeDefault;

  const html = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Sistema de Administración de Condominios GOS</title>' +
    '<style>' +
    'body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;' +
    'background:radial-gradient(circle at 10% 0%, rgba(0,230,160,0.07), transparent 42%),' +
    'radial-gradient(circle at 92% 100%, rgba(0,230,160,0.05), transparent 42%),#05070a;' +
    'font-family:Segoe UI,system-ui,-apple-system,Roboto,Arial,sans-serif;color:#f2f5f7;padding:20px;}' +
    '.box{max-width:420px;text-align:center;background:rgba(255,255,255,0.035);border:1px solid rgba(255,255,255,0.08);' +
    'border-radius:20px;padding:44px 32px;backdrop-filter:blur(24px);}' +
    '.icono{font-size:46px;margin-bottom:18px;}' +
    '.logo{width:90px;height:90px;margin:0 auto 20px;}' +
    '.logo img{width:100%;height:100%;object-fit:contain;filter:drop-shadow(0 6px 20px rgba(0,230,160,0.35));}' +
    'h1{font-size:19px;margin:0 0 12px;}' +
    'p{color:#7c8794;font-size:14px;line-height:1.5;margin:0;}' +
    '</style></head><body>' +
    '<div class="box">' +
    '<div class="logo"><img src="' + LOGO_URL + '" alt="Logo GOS"></div>' +
    '<div class="icono">' + icono + '</div>' +
    '<h1>' + titulo + '</h1>' +
    '<p>' + mensajeFinal + '</p>' +
    '</div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Sistema de Administración de Condominios GOS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/*******************************************************
 * UTILIDAD: indexa encabezados evitando fallos por
 * espacios en blanco invisibles
 *******************************************************/
function _indexarEncabezados(headers) {
  const idx = {};
  headers.forEach(function (h, i) {
    idx[String(h).trim()] = i;
  });
  return idx;
}

function _val(fila, idx, nombreCol) {
  const i = idx[nombreCol];
  return (i === undefined) ? '' : fila[i];
}

/*******************************************************
 * AUTENTICACIÓN
 *******************************************************/
function iniciarSesion(correo, clave) {
  try {
    correo = String(correo || '').trim().toLowerCase();
    clave = String(clave || '').trim();

    if (!correo || !clave) {
      return { ok: false, mensaje: 'Ingrese correo y contraseña.' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const hoja = ss.getSheetByName('PROPIETARIOS');
    if (!hoja) {
      return { ok: false, mensaje: 'No se encontró la hoja "PROPIETARIOS". Ejecute crearHojasYEncabezados().' };
    }

    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) {
      return { ok: false, mensaje: 'La hoja PROPIETARIOS no tiene datos cargados.' };
    }

    const idx = _indexarEncabezados(datos[0]);

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!_val(fila, idx, 'CASA')) continue;

      const correoP = String(_val(fila, idx, 'CORREO PRINCIPAL') || '').trim().toLowerCase();
      const correoS = String(_val(fila, idx, 'CORREO SECUNDARIO') || '').trim().toLowerCase();
      const claveHoja = String(_val(fila, idx, 'CONTRASEÑA') || '').trim();

      if ((correo === correoP || (correoS && correo === correoS)) && clave === claveHoja) {
        return {
          ok: true,
          correo: correo,
          casa: _val(fila, idx, 'CASA'),
          propietarios: _val(fila, idx, 'PROPIETARIOS'),
          rol: String(_val(fila, idx, 'ROL') || 'USUARIO').trim().toUpperCase(),
          estado: _val(fila, idx, 'ESTADO'),
          alicuota: _val(fila, idx, 'ALICUOTA'),
          solvencia: _val(fila, idx, 'SOLVENCIA')
        };
      }
    }
    return { ok: false, mensaje: 'Correo o contraseña incorrectos.' };

  } catch (err) {
    return { ok: false, mensaje: 'Error interno: ' + err.message };
  }
}

/*******************************************************
 * Reverifica en servidor quién es el usuario (por correo)
 * Nunca se confía en el rol que mande el navegador
 *******************************************************/
function _obtenerPropietarioPorCorreo(correo) {
  correo = String(correo || '').trim().toLowerCase();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hoja = ss.getSheetByName('PROPIETARIOS');
  if (!hoja) return null;
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return null;
  const idx = _indexarEncabezados(datos[0]);

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!_val(fila, idx, 'CASA')) continue;
    const cp = String(_val(fila, idx, 'CORREO PRINCIPAL') || '').trim().toLowerCase();
    const cs = String(_val(fila, idx, 'CORREO SECUNDARIO') || '').trim().toLowerCase();
    if (correo === cp || (cs && correo === cs)) {
      return {
        casa: _val(fila, idx, 'CASA'),
        rol: String(_val(fila, idx, 'ROL') || 'USUARIO').trim().toUpperCase(),
        propietarios: _val(fila, idx, 'PROPIETARIOS')
      };
    }
  }
  return null;
}

/*******************************************************
 * Lista de casas (solo disponible para ADMINISTRADOR)
 *******************************************************/
function obtenerListaCasas(correoSesion) {
  const prop = _obtenerPropietarioPorCorreo(correoSesion);
  if (!prop || prop.rol !== 'ADMINISTRADOR') return [];

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hoja = ss.getSheetByName('PROPIETARIOS');
  const datos = hoja.getDataRange().getValues();
  const idx = _indexarEncabezados(datos[0]);

  const lista = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!_val(fila, idx, 'CASA')) continue;
    lista.push({
      casa: _val(fila, idx, 'CASA'),
      propietarios: _val(fila, idx, 'PROPIETARIOS'),
      estado: _val(fila, idx, 'ESTADO')
    });
  }
  lista.sort(function (a, b) { return String(a.casa).localeCompare(String(b.casa)); });
  return lista;
}

/*******************************************************
 * DATOS COMPLETOS DE UNA CASA (sin envoltorio de sesión)
 * Reutilizado por obtenerDatosCasa y generarPdfRelacion
 *******************************************************/
function _datosCasaCompletos(casa) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ---- ESTADO ----
  const hEstado = ss.getSheetByName('ESTADO');
  const dEstado = hEstado.getDataRange().getValues();
  const idxE = _indexarEncabezados(dEstado[0]);
  let estadoCasa = null;
  for (let i = 1; i < dEstado.length; i++) {
    if (String(_val(dEstado[i], idxE, 'CASA')).trim() === String(casa).trim()) {
      estadoCasa = {
        casa: _val(dEstado[i], idxE, 'CASA'),
        propietarios: _val(dEstado[i], idxE, 'PROPIETARIO(S)'),
        recibosVencidos: _val(dEstado[i], idxE, 'RECIBOS VENCIDOS'),
        saldo: _val(dEstado[i], idxE, 'SALDO'),
        estatus: _val(dEstado[i], idxE, 'ESTATUS')
      };
      break;
    }
  }

  // ---- INFO PROPIETARIO ----
  const hProp = ss.getSheetByName('PROPIETARIOS');
  const dProp = hProp.getDataRange().getValues();
  const idxP = _indexarEncabezados(dProp[0]);
  let infoCasa = null;
  for (let i = 1; i < dProp.length; i++) {
    if (String(_val(dProp[i], idxP, 'CASA')).trim() === String(casa).trim()) {
      infoCasa = {
        casa: _val(dProp[i], idxP, 'CASA'),
        propietarios: _val(dProp[i], idxP, 'PROPIETARIOS'),
        calle: _val(dProp[i], idxP, 'CALLE'),
        alicuota: _val(dProp[i], idxP, 'ALICUOTA'),
        solvencia: _val(dProp[i], idxP, 'SOLVENCIA'),
        estado: _val(dProp[i], idxP, 'ESTADO')
      };
      break;
    }
  }

  // ---- CARGOS ----
  const hCargos = ss.getSheetByName('CARGOS');
  const dCargos = hCargos.getDataRange().getValues();
  const idxC = _indexarEncabezados(dCargos[0]);
  const cargos = [];
  for (let i = 1; i < dCargos.length; i++) {
    if (String(_val(dCargos[i], idxC, 'CASA')).trim() === String(casa).trim()) {
      const referencia = _val(dCargos[i], idxC, 'REFERENCIA');
      cargos.push({
        referencia: referencia,
        monto: Number(_val(dCargos[i], idxC, 'MONTO ($)')) || 0,
        fecha: _inferirFechaCargo(referencia)
      });
    }
  }

  // ---- PAGOS ----
  const hPagos = ss.getSheetByName('PAGOS');
  const dPagos = hPagos.getDataRange().getValues();
  const idxPa = _indexarEncabezados(dPagos[0]);
  const pagos = [];
  for (let i = 1; i < dPagos.length; i++) {
    if (String(_val(dPagos[i], idxPa, 'CASA')).trim() === String(casa).trim()) {
      pagos.push({
        fecha: _val(dPagos[i], idxPa, 'FECHA'),
        descripcion: _val(dPagos[i], idxPa, 'DESCRIPCION') || 'PAGO',
        montoBs: Number(_val(dPagos[i], idxPa, 'MONTO (BS)')) || 0,
        refUsd: _parsearMonedaTexto(_val(dPagos[i], idxPa, 'Ref. US $'))
      });
    }
  }

  return {
    estado: estadoCasa,
    info: infoCasa,
    relacion: _construirRelacion(cargos, pagos)
  };
}

/*******************************************************
 * Datos completos de una casa: estado, info y relación
 * (envuelve _datosCasaCompletos con validación de sesión)
 *******************************************************/
function obtenerDatosCasa(correoSesion, casaSolicitada) {
  const prop = _obtenerPropietarioPorCorreo(correoSesion);
  if (!prop) return { ok: false, mensaje: 'Sesión inválida. Vuelva a iniciar sesión.' };

  let casa = casaSolicitada;
  if (prop.rol !== 'ADMINISTRADOR') {
    casa = prop.casa; // Un USUARIO solo puede ver su propia casa, sin excepción
  }
  if (!casa) casa = prop.casa;

  const datos = _datosCasaCompletos(casa);

  return {
    ok: true,
    rolSesion: prop.rol,
    casa: casa,
    estado: datos.estado,
    info: datos.info,
    relacion: datos.relacion
  };
}

/*******************************************************
 * UTILIDADES
 *******************************************************/
function _parsearMonedaTexto(txt) {
  if (txt === null || txt === undefined || txt === '') return 0;
  if (typeof txt === 'number') return txt;
  const limpio = String(txt).replace(/[^0-9.\-]/g, '');
  const n = parseFloat(limpio);
  return isNaN(n) ? 0 : n;
}

function _inferirFechaCargo(referencia) {
  const ref = String(referencia || '').toUpperCase();
  if (ref.indexOf('SALDO') >= 0) {
    return new Date(new Date().getFullYear(), 0, 1);
  }
  for (const mes in MESES) {
    if (ref.indexOf(mes) >= 0) {
      return new Date(new Date().getFullYear(), MESES[mes] - 1, 1);
    }
  }
  return null;
}

function _construirRelacion(cargos, pagos) {
  const items = [];

  cargos.forEach(function (c) {
    const esSaldoInicial = String(c.referencia || '').toUpperCase().indexOf('SALDO') >= 0;
    items.push({
      tipo: esSaldoInicial ? 'SALDO_INICIAL' : 'CARGO',
      fecha: c.fecha,
      descripcion: c.referencia,
      monto: null,
      ref: esSaldoInicial ? null : c.monto,
      saldoInicial: esSaldoInicial ? c.monto : null,
      ordenSinFecha: c.fecha ? 0 : 1
    });
  });

  pagos.forEach(function (p) {
    items.push({
      tipo: 'PAGO',
      fecha: p.fecha,
      descripcion: p.descripcion,
      monto: p.montoBs,
      ref: -Math.abs(p.refUsd),
      saldoInicial: null,
      ordenSinFecha: 0
    });
  });

  items.sort(function (a, b) {
    if (a.ordenSinFecha !== b.ordenSinFecha) return a.ordenSinFecha - b.ordenSinFecha;
    if (!a.fecha && !b.fecha) return 0;
    if (!a.fecha) return 1;
    if (!b.fecha) return -1;
    const diff = new Date(a.fecha) - new Date(b.fecha);
    if (diff !== 0) return diff;
    const orden = { SALDO_INICIAL: 0, CARGO: 0, PAGO: 1 };
    return orden[a.tipo] - orden[b.tipo];
  });

  let saldoAcumulado = 0;
  return items.map(function (it) {
    const esPagoEspecial = it.tipo === 'PAGO' &&
      String(it.descripcion || '').trim().toUpperCase() === 'PAGO CUOTA ESPECIAL';

    if (it.tipo === 'SALDO_INICIAL') {
      saldoAcumulado = it.saldoInicial;
    } else if (!esPagoEspecial) {
      saldoAcumulado += it.ref;
    }

    return {
      fecha: it.fecha ? Utilities.formatDate(new Date(it.fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '—',
      descripcion: it.descripcion,
      monto: it.monto,
      ref: it.tipo === 'SALDO_INICIAL' ? null : it.ref,
      saldo: esPagoEspecial ? null : Number(saldoAcumulado.toFixed(2)),
      tipo: it.tipo
    };
  });
}

/*******************************************************
 * CAMBIAR CONTRASEÑA
 * Verifica la contraseña actual antes de actualizar
 *******************************************************/
function cambiarContrasena(correoSesion, claveActual, claveNueva) {
  try {
    correoSesion = String(correoSesion || '').trim().toLowerCase();
    claveActual = String(claveActual || '').trim();
    claveNueva = String(claveNueva || '').trim();

    if (!correoSesion) return { ok: false, mensaje: 'Sesión inválida.' };
    if (!claveActual || !claveNueva) return { ok: false, mensaje: 'Complete todos los campos.' };
    if (claveNueva.length < 4) return { ok: false, mensaje: 'La nueva contraseña debe tener al menos 4 caracteres.' };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const hoja = ss.getSheetByName('PROPIETARIOS');
    const datos = hoja.getDataRange().getValues();
    const idx = _indexarEncabezados(datos[0]);
    const colClave = idx['CONTRASEÑA'];

    if (colClave === undefined) {
      return { ok: false, mensaje: 'No se encontró la columna CONTRASEÑA en la hoja.' };
    }

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const cp = String(_val(fila, idx, 'CORREO PRINCIPAL') || '').trim().toLowerCase();
      const cs = String(_val(fila, idx, 'CORREO SECUNDARIO') || '').trim().toLowerCase();

      if (correoSesion === cp || (cs && correoSesion === cs)) {
        const claveHoja = String(_val(fila, idx, 'CONTRASEÑA') || '').trim();
        if (claveHoja !== claveActual) {
          return { ok: false, mensaje: 'La contraseña actual no es correcta.' };
        }
        hoja.getRange(i + 1, colClave + 1).setValue(claveNueva);
        return { ok: true, mensaje: 'Contraseña actualizada correctamente.' };
      }
    }
    return { ok: false, mensaje: 'Usuario no encontrado.' };

  } catch (err) {
    return { ok: false, mensaje: 'Error interno: ' + err.message };
  }
}

/*******************************************************
 * GENERAR PDF DE LA RELACIÓN DE PAGOS Y CARGOS
 *******************************************************/
function generarPdfRelacion(correoSesion, casaSolicitada) {
  let ssTemp = null;
  try {
    const prop = _obtenerPropietarioPorCorreo(correoSesion);
    if (!prop) return { ok: false, mensaje: 'Sesión inválida.' };

    let casa = casaSolicitada;
    if (prop.rol !== 'ADMINISTRADOR') casa = prop.casa;
    if (!casa) casa = prop.casa;

    const datos = _datosCasaCompletos(casa);
    const info = datos.info || {};
    const estado = datos.estado || {};
    const relacion = datos.relacion || [];

    ssTemp = SpreadsheetApp.create('PDF_TEMP_' + casa + '_' + new Date().getTime());
    const hoja = ssTemp.getSheets()[0];
    hoja.setName('Relacion');

    // Anchos de columna fijos (evita el espacio gigante entre columnas)
    hoja.setColumnWidth(1, 150);  // Fecha / etiquetas
    hoja.setColumnWidth(2, 260);  // Descripción
    hoja.setColumnWidth(3, 90);   // Monto (Bs)
    hoja.setColumnWidth(4, 80);   // Ref ($)
    hoja.setColumnWidth(5, 90);   // Saldo

    hoja.setRowHeight(1, 46);
    hoja.setRowHeight(2, 24);

    // ---- Encabezado con título ----
    hoja.getRange(1, 1, 1, 4).merge()
      .setValue('SISTEMA DE ADMINISTRACIÓN INTEGRAL DE CONDOMINIOS')
      .setFontSize(13).setFontWeight('bold').setVerticalAlignment('middle');
    hoja.getRange(2, 1, 1, 4).merge()
      .setValue('RELACIÓN DE PAGOS Y CARGOS')
      .setFontSize(10).setFontColor('#555555').setVerticalAlignment('middle');

    // ---- Logo arriba a la derecha ----
    try {
      const logoBlob = UrlFetchApp.fetch(LOGO_URL).getBlob();
      hoja.insertImage(logoBlob, 5, 1).setWidth(55).setHeight(55);
    } catch (eLogo) {
      // Si falla la carga del logo, el PDF se genera igual sin él
    }

    let fila = 4;
    hoja.getRange(fila, 1).setValue('Casa:').setFontWeight('bold');
    hoja.getRange(fila, 2).setValue(casa).setHorizontalAlignment('left');
    fila++;
    hoja.getRange(fila, 1).setValue('Propietario(s):').setFontWeight('bold');
    hoja.getRange(fila, 2).setValue(info.propietarios || '').setHorizontalAlignment('left');
    fila++;
    hoja.getRange(fila, 1).setValue('Saldo actual:').setFontWeight('bold');
    hoja.getRange(fila, 2).setValue(estado.saldo || '').setHorizontalAlignment('left');
    fila++;
    hoja.getRange(fila, 1).setValue('Recibos vencidos:').setFontWeight('bold');
    hoja.getRange(fila, 2).setValue(estado.recibosVencidos !== undefined ? estado.recibosVencidos : '').setHorizontalAlignment('left');
    fila += 2;

    const headers = ['Fecha', 'Descripción', 'Monto (Bs)', 'Ref ($)', 'Saldo'];
    hoja.getRange(fila, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    fila++;

    if (relacion.length > 0) {
      const filasDatos = relacion.map(function (m) {
        return [
          m.fecha,
          m.descripcion,
          (m.monto !== null && m.monto !== undefined) ? m.monto : '',
          (m.ref !== null && m.ref !== undefined) ? m.ref : '',
          m.saldo
        ];
      });
      hoja.getRange(fila, 1, filasDatos.length, headers.length).setValues(filasDatos);
    }

    SpreadsheetApp.flush();

    const token = ScriptApp.getOAuthToken();
    const url = 'https://docs.google.com/spreadsheets/d/' + ssTemp.getId() +
      '/export?exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true' +
      '&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false' +
      '&gid=' + hoja.getSheetId();

    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    const blob = response.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    DriveApp.getFileById(ssTemp.getId()).setTrashed(true);

    const nombreArchivo = 'Relacion_' + String(casa).replace(/[^a-zA-Z0-9_-]/g, '_') + '.pdf';
    return { ok: true, base64: base64, filename: nombreArchivo };

  } catch (err) {
    if (ssTemp) {
      try { DriveApp.getFileById(ssTemp.getId()).setTrashed(true); } catch (e2) {}
    }
    return { ok: false, mensaje: 'Error generando PDF: ' + err.message };
  }
}

/*******************************************************
 * Construye una fila de la tabla de datos del correo
 *******************************************************/
function _filaDatoCorreo(label, valor) {
  return '<tr>' +
    '<td style="padding:12px 16px;border-bottom:1px solid #232838;color:#7c8794;font-size:11px;text-transform:uppercase;letter-spacing:.5px;width:40%;">' + label + '</td>' +
    '<td style="padding:12px 16px;border-bottom:1px solid #232838;color:#f2f5f7;font-size:14px;font-weight:700;">' + valor + '</td>' +
    '</tr>';
}

/*******************************************************
 * Correo elegante de confirmación para el propietario
 *******************************************************/
function _construirCorreoConfirmacionPago(casa, propietario, fecha, monto, entidadBancaria) {
  return '<div style="background:#0d0f14;padding:30px 0;font-family:Arial,Helvetica,sans-serif;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">' +
    '<table role="presentation" width="480" cellpadding="0" cellspacing="0" style="background:#12151c;border-radius:16px;overflow:hidden;border:1px solid #232838;">' +
      '<tr><td style="background:linear-gradient(135deg,#00e6a0,#00b894);padding:28px 30px;text-align:center;">' +
        '<img src="' + LOGO_URL + '" width="60" height="60" style="display:block;margin:0 auto 10px;border-radius:14px;" alt="Logo GOS">' +
        '<div style="color:#04140d;font-size:16px;font-weight:800;letter-spacing:.5px;">SISTEMA DE ADMINISTRACIÓN DE CONDOMINIOS GOS</div>' +
      '</td></tr>' +
      '<tr><td style="padding:30px 30px 10px;">' +
        '<h2 style="margin:0 0 14px;color:#f2f5f7;font-size:18px;">¡Hemos recibido tu pago!</h2>' +
        '<p style="margin:0 0 18px;color:#aeb6c2;font-size:14px;line-height:1.6;">' +
        'Se informa que hemos recibido el pago, el cual se encuentra en proceso de verificación. ' +
        'Una vez confirmado, se le estará enviando un comprobante de pago al correo electrónico registrado.' +
        '</p>' +
      '</td></tr>' +
      '<tr><td style="padding:0 30px 26px;">' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#1a1e28;border-radius:12px;border:1px solid #232838;">' +
          _filaDatoCorreo('Casa', casa) +
          _filaDatoCorreo('Propietario', propietario) +
          _filaDatoCorreo('Fecha del pago', fecha) +
          _filaDatoCorreo('Monto', monto) +
          _filaDatoCorreo('Entidad bancaria', entidadBancaria) +
        '</table>' +
      '</td></tr>' +
      '<tr><td style="padding:0 30px 30px;text-align:center;">' +
        '<div style="display:inline-block;background:#00e6a01f;border:1px solid #00e6a055;color:#00e6a0;font-size:12px;font-weight:700;padding:8px 16px;border-radius:20px;letter-spacing:.4px;">EN VERIFICACIÓN</div>' +
      '</td></tr>' +
      '<tr><td style="padding:18px 30px;background:#0d0f14;border-top:1px solid #232838;text-align:center;">' +
        '<p style="margin:0;color:#5b6472;font-size:11px;">Este es un mensaje automático, no responda a este correo.</p>' +
      '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
  '</div>';
}

/*******************************************************
 * REGISTRAR PAGO (con imagen adjunta por correo)
 * Envía un correo interno al administrador y un correo
 * elegante de confirmación al propietario.
 *******************************************************/
function registrarPago(correoSesion, casaSolicitada, fecha, monto, entidadBancaria, observacion, imagenBase64, imagenNombre, imagenTipo) {
  try {
    const prop = _obtenerPropietarioPorCorreo(correoSesion);
    if (!prop) return { ok: false, mensaje: 'Sesión inválida.' };

    let casa = casaSolicitada;
    if (prop.rol !== 'ADMINISTRADOR') casa = prop.casa;
    if (!casa) casa = prop.casa;

    fecha = String(fecha || '').trim();
    monto = String(monto || '').trim();
    entidadBancaria = String(entidadBancaria || '').trim();
    observacion = String(observacion || '').trim();

    if (!fecha || !monto || !entidadBancaria) {
      return { ok: false, mensaje: 'Complete todos los campos.' };
    }

    let blobAdjunto = null;
    if (imagenBase64 && imagenNombre) {
      const bytes = Utilities.base64Decode(imagenBase64);
      blobAdjunto = Utilities.newBlob(bytes, imagenTipo || 'image/jpeg', imagenNombre);
    }

    // ---- Correo interno para el administrador ----
    const asuntoAdmin = 'Registro de pago - Casa ' + casa;
    const cuerpoAdmin = '<div style="font-family:Arial,sans-serif;">' +
      '<h2 style="color:#1a1a2e;">Nuevo registro de pago</h2>' +
      '<p><b>Casa:</b> ' + casa + '</p>' +
      '<p><b>Propietario:</b> ' + (prop.propietarios || '') + '</p>' +
      '<p><b>Correo del propietario:</b> ' + correoSesion + '</p>' +
      '<p><b>Fecha del pago:</b> ' + fecha + '</p>' +
      '<p><b>Monto:</b> ' + monto + '</p>' +
      '<p><b>Entidad bancaria:</b> ' + entidadBancaria + '</p>' +
      (observacion ? '<p><b>Observación:</b> ' + observacion + '</p>' : '') +
      '</div>';

    const opcionesAdmin = { htmlBody: cuerpoAdmin, name: 'Sistema de Administración de Condominios GOS' };
    if (blobAdjunto) opcionesAdmin.attachments = [blobAdjunto];
    MailApp.sendEmail('gos.condominio@gmail.com', asuntoAdmin, 'Se registró un nuevo pago. Ver detalles en el correo.', opcionesAdmin);

    // ---- Correo elegante de confirmación para el propietario ----
    try {
      const asuntoPropietario = 'Confirmación de registro de pago - Casa ' + casa;
      const cuerpoPropietario = _construirCorreoConfirmacionPago(casa, prop.propietarios || '', fecha, monto, entidadBancaria, observacion);
      const opcionesPropietario = { htmlBody: cuerpoPropietario, name: 'Sistema de Administración de Condominios GOS' };
      if (blobAdjunto) opcionesPropietario.attachments = [blobAdjunto];
      MailApp.sendEmail(correoSesion, asuntoPropietario, 'Hemos recibido tu pago y está en proceso de verificación.', opcionesPropietario);
    } catch (eCorreoProp) {
      // Si falla el correo de confirmación, no se interrumpe el registro
    }

    // Registro opcional en hoja para auditoría (se crea sola la primera vez)
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let hoja = ss.getSheetByName('REPORTES_PAGO');
      if (!hoja) {
        hoja = ss.insertSheet('REPORTES_PAGO');
        hoja.getRange(1, 1, 1, 7).setValues([['FECHA REGISTRO', 'CASA', 'PROPIETARIO', 'FECHA PAGO', 'MONTO', 'ENTIDAD', 'OBSERVACION']])
          .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
        hoja.setFrozenRows(1);
      }
      hoja.appendRow([new Date(), casa, prop.propietarios || '', fecha, monto, entidadBancaria, observacion]);
    } catch (eLog) {}

    return { ok: true, mensaje: 'Pago registrado correctamente. Te enviamos un correo de confirmación.' };

  } catch (err) {
    return { ok: false, mensaje: 'Error registrando el pago: ' + err.message };
  }
}