const COL_CODIGO = 'Invite_Unique_Code';
const HOJA_LOG = 'Log';

const PUNTOS = {
  'Palermo': 'Av. Corrientes y Scalabrini Ortiz · 19:50 hs',
  'Recoleta': 'Recoleta · 20:15 hs'
};

const EDAD_ADULTA = 'Adulto';

const LARGO_MAXIMO = 500;

const COLUMNAS = [
  COL_CODIGO, 'Name_DB', 'Surname_DB', 'Age_Range', 'Transfer_Site',
  'Timestamp', 'Assistance_Confirmation', 'Name_Invite', 'Surname_Invite',
  'Phone', 'Adult_phone', 'Transfer_use_outbound', 'Transfer_use_inbound',
  'Dietary_restrictions', 'Dietary_restrictions_other', 'Song_preferece'
];

const ASISTE = 'Sí, ahí voy a estar';

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.action === 'lookup') return json(lookup(params.code));
  return json({ ok: true, mensaje: 'RSVP de Clari XV activo.' });
}

function doPost(e) {
  let datos;
  try {
    datos = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: 'formato' });
  }
  if (datos.action === 'lookup') return json(lookup(datos.code));
  return json(confirmar(datos));
}

function lookup(codigo) {
  const ctx = ubicar(codigo);
  if (ctx.error) return { ok: false, error: ctx.error, falta: ctx.falta };
  if (ctx.usado) return { ok: false, error: 'usado', fecha: ctx.fecha };
  return {
    ok: true,
    code: ctx.codigo,
    nombre: ctx.nombre,
    apellido: ctx.apellido,
    menor: ctx.menor,
    punto: ctx.punto
  };
}

function confirmar(datos) {
  const codigo = normalizar(datos && datos.code);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'ocupado' };
  try {
    const ctx = ubicar(codigo);
    if (ctx.error) return registrar(codigo, ctx.error, datos, { ok: false, error: ctx.error, falta: ctx.falta });
    if (ctx.usado) return registrar(codigo, 'usado', datos, { ok: false, error: 'usado', fecha: ctx.fecha });

    const respuesta = armar(datos, ctx);
    if (respuesta.campo) return registrar(codigo, 'incompleto:' + respuesta.campo, datos, { ok: false, error: 'incompleto', campo: respuesta.campo });

    escribir(ctx, respuesta);
    return registrar(codigo, 'ok', datos, { ok: true });
  } catch (err) {
    return registrar(codigo, 'error: ' + err.message, datos, { ok: false, error: 'error' });
  } finally {
    lock.releaseLock();
  }
}

function ubicar(codigo) {
  const code = normalizar(codigo);
  if (code.length !== 6) return { error: 'invalido' };

  const hoja = hojaInvitados();
  if (!hoja) return { error: 'config', falta: [COL_CODIGO] };

  const valores = hoja.getDataRange().getValues();
  const idx = indices(valores[0]);
  const falta = COLUMNAS.filter(function (nombre) { return idx[nombre] === undefined; });
  if (falta.length) return { error: 'config', falta: falta };

  const filas = [];
  for (let i = 1; i < valores.length; i++) {
    if (normalizar(valores[i][idx[COL_CODIGO]]) === code) filas.push(i);
  }
  if (!filas.length) return { error: 'invalido' };
  if (filas.length > 1) return { error: 'duplicado' };

  const fila = valores[filas[0]];
  const sitio = String(fila[idx.Transfer_Site] || '').trim();
  const marca = fila[idx.Timestamp];

  return {
    hoja: hoja,
    idx: idx,
    numeroFila: filas[0] + 1,
    codigo: code,
    nombre: String(fila[idx.Name_DB] || '').trim(),
    apellido: limpiarApellido(fila[idx.Surname_DB]),
    menor: !esAdulto(fila[idx.Age_Range]),
    sitio: sitio,
    punto: PUNTOS[sitio] || '',
    usado: marca !== '' && marca !== null,
    fecha: marca ? Utilities.formatDate(new Date(marca), 'America/Argentina/Buenos_Aires', 'd/M/yyyy') : ''
  };
}

function armar(datos, ctx) {
  const d = datos || {};
  const nombre = texto(d.nombre) || ctx.nombre;
  const apellido = texto(d.apellido) || ctx.apellido;
  const telefono = normalizarTel(d.telefono);
  const adultoTel = normalizarTel(d.adultoTel);
  const restriccion = texto(d.restriccion);
  const restriccionOtra = restriccion === 'Otra' ? texto(d.restriccionOtra) : '';
  const cancion = texto(d.cancion);
  const ofrece = !!ctx.punto;
  const transfer = ofrece ? texto(d.transfer) : '';
  const usaTransfer = transfer === 'Sí';
  const contesto = transfer === 'Sí' || transfer === 'No';

  const campos = {
    nombre: nombre,
    apellido: apellido,
    telefono: telefono,
    adultoTel: ctx.menor ? adultoTel : 'n/a',
    restriccion: restriccion
  };
  for (const campo in campos) {
    if (!campos[campo]) return { campo: campo };
  }

  const textos = [nombre, apellido, telefono, adultoTel, restriccion, restriccionOtra, cancion];
  for (let i = 0; i < textos.length; i++) {
    if (textos[i].length > LARGO_MAXIMO) return { campo: 'largo' };
  }

  return {
    nombre: nombre,
    apellido: apellido,
    telefono: telefono,
    adultoTel: ctx.menor ? adultoTel : '',
    restriccion: restriccion,
    restriccionOtra: restriccionOtra,
    ida: contesto ? (usaTransfer ? texto(d.ida) : 'No') : '',
    vuelta: contesto ? (usaTransfer ? texto(d.vuelta) : 'No') : '',
    cancion: cancion
  };
}

function escribir(ctx, r) {
  const hoja = ctx.hoja;
  const idx = ctx.idx;
  const fila = ctx.numeroFila;

  const celda = function (columna, valor) {
    hoja.getRange(fila, idx[columna] + 1).setValue(valor);
  };

  const celdaTexto = function (columna, valor) {
    hoja.getRange(fila, idx[columna] + 1).setNumberFormat('@').setValue(valor);
  };

  celda('Assistance_Confirmation', ASISTE);
  celda('Name_Invite', r.nombre);
  celda('Surname_Invite', r.apellido);
  celdaTexto('Phone', r.telefono);
  celdaTexto('Adult_phone', r.adultoTel);
  celda('Transfer_use_outbound', r.ida);
  celda('Transfer_use_inbound', r.vuelta);
  celda('Dietary_restrictions', r.restriccion);
  celda('Dietary_restrictions_other', r.restriccionOtra);
  celda('Song_preferece', r.cancion);
  celda('Timestamp', new Date());
}

function hojaInvitados() {
  const hojas = SpreadsheetApp.getActive().getSheets();
  for (let i = 0; i < hojas.length; i++) {
    const encabezados = hojas[i].getRange(1, 1, 1, hojas[i].getLastColumn() || 1).getValues()[0];
    if (encabezados.indexOf(COL_CODIGO) !== -1) return hojas[i];
  }
  return null;
}

function indices(encabezados) {
  const mapa = {};
  encabezados.forEach(function (nombre, i) {
    const clave = String(nombre || '').trim();
    if (clave && mapa[clave] === undefined) mapa[clave] = i;
  });
  return mapa;
}

function registrar(codigo, resultado, datos, respuesta) {
  try {
    const libro = SpreadsheetApp.getActive();
    let hoja = libro.getSheetByName(HOJA_LOG);
    if (!hoja) {
      hoja = libro.insertSheet(HOJA_LOG);
      hoja.appendRow(['Fecha', 'Codigo', 'Resultado', 'Datos']);
    }
    hoja.appendRow([new Date(), codigo, resultado, JSON.stringify(datos || {})]);
  } catch (err) {}
  return respuesta;
}

function normalizarTel(valor) {
  const crudo = texto(valor);
  return crudo.indexOf('+') === 0 ? '+' + digitos(crudo) : digitos(crudo);
}

function esAdulto(valor) {
  return texto(valor).toLowerCase() === EDAD_ADULTA.toLowerCase();
}

function limpiarApellido(valor) {
  const t = String(valor || '').trim();
  return /[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]/.test(t) ? t : '';
}

function normalizar(valor) {
  return String(valor == null ? '' : valor).toUpperCase().replace(/[^A-Z]/g, '');
}

function texto(valor) {
  return String(valor == null ? '' : valor).trim();
}

function digitos(valor) {
  return String(valor == null ? '' : valor).replace(/\D/g, '');
}

function json(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
