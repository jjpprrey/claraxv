const COL_CODIGO = 'Invite_Unique_Code';
const HOJA_LOG = 'Log';

const PUNTOS = {
  'Palermo': 'Av. Corrientes y Scalabrini Ortiz · 19:50 hs',
  'Recoleta': 'Recoleta · 20:15 hs'
};

const EDAD_ADULTA = 'Adulto';

const LARGO_MAXIMO = 500;

const COL_RESPUESTA = [
  'Assistance_Confirmation', 'Name_Invite', 'Surname_Invite', 'Phone', 'Adult_phone',
  'Transfer_use_outbound', 'Transfer_use_inbound',
  'Dietary_restrictions', 'Dietary_restrictions_other', 'Song_preferece'
];

// Sheets se come el + y reformatea los números largos si la celda no es texto.
const COL_LITERAL = ['Phone', 'Adult_phone'];

const COLUMNAS = [
  COL_CODIGO, 'Name_DB', 'Surname_DB', 'Age_Range', 'Transfer_Site', 'Timestamp'
].concat(COL_RESPUESTA);

const ASISTE = 'Sí, ahí voy a estar';
const NO_ASISTE = 'No voy a poder ir';

const ACCION = { nuevo: 'alta', confirmado: 'modificacion', cancelado: 'reconfirmacion' };

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
  if (datos.action === 'cancelar') return json(cancelar(datos));
  return json(confirmar(datos));
}

// Para un disparador cada 5 minutos: evita que la primera visita del día
// pague el arranque en frío del proyecto.
function calentar() {
  ubicar('AAAAAA');
}

function lookup(codigo) {
  const ctx = ubicar(codigo);
  if (ctx.error) return { ok: false, error: ctx.error, falta: ctx.falta };
  return {
    ok: true,
    code: ctx.codigo,
    nombre: ctx.nombre,
    apellido: ctx.apellido,
    menor: ctx.menor,
    punto: ctx.punto,
    estado: ctx.estado,
    fecha: ctx.fecha,
    respuestas: respuestas(ctx.previo)
  };
}

function confirmar(datos) {
  const codigo = normalizar(datos && datos.code);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'ocupado' };
  try {
    const ctx = ubicar(codigo);
    if (ctx.error) return registrar(codigo, 'rechazo:' + ctx.error, {}, datos, { ok: false, error: ctx.error, falta: ctx.falta });

    const respuesta = armar(datos, ctx);
    if (respuesta.campo) return registrar(codigo, 'rechazo:incompleto:' + respuesta.campo, ctx.previo, datos, { ok: false, error: 'incompleto', campo: respuesta.campo });

    const escrito = escribir(ctx, respuesta);
    return registrar(codigo, ACCION[ctx.estado], ctx.previo, escrito, { ok: true, estado: 'confirmado', respuestas: respuestas(escrito) });
  } catch (err) {
    return registrar(codigo, 'error: ' + err.message, {}, datos, { ok: false, error: 'error' });
  } finally {
    lock.releaseLock();
  }
}

function cancelar(datos) {
  const codigo = normalizar(datos && datos.code);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'ocupado' };
  try {
    const ctx = ubicar(codigo);
    if (ctx.error) return registrar(codigo, 'rechazo:' + ctx.error, {}, datos, { ok: false, error: ctx.error, falta: ctx.falta });
    if (ctx.estado === 'cancelado') return { ok: true, estado: 'cancelado', respuestas: respuestas(ctx.previo) };

    const borrado = borrar(ctx);
    return registrar(codigo, 'cancelacion', ctx.previo, borrado, { ok: true, estado: 'cancelado', respuestas: respuestas(borrado) });
  } catch (err) {
    return registrar(codigo, 'error: ' + err.message, {}, datos, { ok: false, error: 'error' });
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

  const previo = {};
  COL_RESPUESTA.forEach(function (nombre) { previo[nombre] = texto(fila[idx[nombre]]); });

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
    previo: previo,
    estado: estado(marca, previo.Assistance_Confirmation),
    fecha: marca ? Utilities.formatDate(new Date(marca), 'America/Argentina/Buenos_Aires', 'd/M/yyyy') : ''
  };
}

function estado(marca, asistencia) {
  if (marca === '' || marca === null || marca === undefined) return 'nuevo';
  return texto(asistencia).toLowerCase() === NO_ASISTE.toLowerCase() ? 'cancelado' : 'confirmado';
}

// Lo que la fila tiene guardado, con los nombres que usa el formulario.
function respuestas(previo) {
  const ida = texto(previo.Transfer_use_outbound);
  const vuelta = texto(previo.Transfer_use_inbound);
  return {
    nombre: texto(previo.Name_Invite),
    apellido: texto(previo.Surname_Invite),
    telefono: texto(previo.Phone),
    adultoTel: texto(previo.Adult_phone),
    restriccion: texto(previo.Dietary_restrictions),
    restriccionOtra: texto(previo.Dietary_restrictions_other),
    transfer: (ida || vuelta) ? ((ida === 'Sí' || vuelta === 'Sí') ? 'Sí' : 'No') : '',
    ida: ida,
    vuelta: vuelta,
    cancion: texto(previo.Song_preferece)
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
  const escrito = {
    Assistance_Confirmation: ASISTE,
    Name_Invite: r.nombre,
    Surname_Invite: r.apellido,
    Phone: r.telefono,
    Adult_phone: r.adultoTel,
    Transfer_use_outbound: r.ida,
    Transfer_use_inbound: r.vuelta,
    Dietary_restrictions: r.restriccion,
    Dietary_restrictions_other: r.restriccionOtra,
    Song_preferece: r.cancion
  };
  volcar(ctx, escrito);
  return escrito;
}

function borrar(ctx) {
  const escrito = {};
  COL_RESPUESTA.forEach(function (nombre) { escrito[nombre] = ''; });
  escrito.Assistance_Confirmation = NO_ASISTE;
  volcar(ctx, escrito);
  return escrito;
}

function volcar(ctx, valores) {
  const hoja = ctx.hoja;
  const idx = ctx.idx;
  const fila = ctx.numeroFila;

  COL_RESPUESTA.forEach(function (nombre) {
    const celda = hoja.getRange(fila, idx[nombre] + 1);
    if (COL_LITERAL.indexOf(nombre) !== -1) celda.setNumberFormat('@');
    celda.setValue(valores[nombre]);
  });
  hoja.getRange(fila, idx.Timestamp + 1).setValue(new Date());
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

function registrar(codigo, accion, antes, despues, respuesta) {
  try {
    const libro = SpreadsheetApp.getActive();
    let hoja = libro.getSheetByName(HOJA_LOG);
    if (!hoja) hoja = libro.insertSheet(HOJA_LOG);
    if (!hoja.getLastRow()) hoja.appendRow(['Fecha', 'Codigo', 'Accion', 'Antes', 'Despues']);
    hoja.appendRow([new Date(), codigo, accion, JSON.stringify(antes || {}), JSON.stringify(despues || {})]);
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
