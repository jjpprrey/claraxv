const fs = require('fs');
const assert = require('assert');

const CABECERA = [
  'ID_Guest', 'Invite_Unique_Code', 'Invite_Link', 'Name_DB', 'Surname_DB', 'Age_Range',
  'Guest_Host_Relation', 'Transfer_Site', 'Timestamp', 'Assistance_Confirmation',
  'Name_Invite', 'Surname_Invite', 'Phone', 'Adult_phone', 'Transfer_use_outbound',
  'Transfer_use_inbound', 'Dietary_restrictions', 'Dietary_restrictions_other', 'Song_preferece'
];

function filaVacia(id, code, nombre, apellido, edad, sitio) {
  const f = new Array(CABECERA.length).fill('');
  f[0] = id; f[1] = code; f[3] = nombre; f[4] = apellido; f[5] = edad; f[7] = sitio;
  return f;
}

function hojaFalsa(nombre, datos) {
  return {
    nombre: nombre,
    datos: datos,
    getLastColumn: () => datos[0].length,
    getDataRange: () => ({ getValues: () => datos }),
    getRange: (f, c, nf, nc) => {
      const rango = {
        getValues: () => datos.slice(f - 1, f - 1 + (nf || 1)).map(r => r.slice(c - 1, c - 1 + (nc || 1))),
        setValue: v => { datos[f - 1][c - 1] = v; return rango; },
        setNumberFormat: () => rango
      };
      return rango;
    },
    getLastRow: () => datos.length,
    appendRow: r => datos.push(r)
  };
}

function entorno(filas) {
  const invitados = hojaFalsa('Lista Invitados', [CABECERA.slice()].concat(filas));
  const libro = {
    hojas: [invitados],
    getSheets() { return this.hojas; },
    getSheetByName(n) { return this.hojas.find(h => h.nombre === n) || null; },
    insertSheet(n) { const h = hojaFalsa(n, []); this.hojas.push(h); return h; }
  };
  global.SpreadsheetApp = { getActive: () => libro };
  global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  global.Utilities = { formatDate: d => new Date(d).toLocaleDateString('es-AR') };
  global.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: t => ({ setMimeType: () => t }) };
  return { libro, invitados };
}

const src = fs.readFileSync(__dirname + '/Codigo.gs', 'utf8');
eval(src + '\nglobal.__api = { lookup, confirmar, cancelar, ubicar };');
const api = global.__api;

const base = () => [
  filaVacia(3, 'KXZKUQ', 'Clara', 'Herrera', 'Adulto', 'Recoleta'),
  filaVacia(7, 'PPMUMN', 'Julia', 'Armani', 'Menor 18', 'Recoleta'),
  filaVacia(61, 'QNNTVJ', 'Amiga de Carmen', '', 'Adulto', ''),
  filaVacia(30, 'XUBPGE', 'Felix o Margarita', '—', 'Adulto', 'Recoleta'),
  filaVacia(78, 'HDKBIS', 'Maitena', 'Undank', 'Menor 18', 'Palermo'),
  filaVacia(200, 'WQWDEE', 'Invitado', 'Nuevo', '', 'Recoleta')
];

// lookup
let e = entorno(base());
let r = api.lookup('kxzkuq');
assert.deepStrictEqual([r.ok, r.nombre, r.apellido, r.menor, r.punto],
  [true, 'Clara', 'Herrera', false, 'Recoleta · 20:15 hs'], 'lookup adulto Recoleta');

assert.strictEqual(api.lookup('HDKBIS').punto, 'Av. Corrientes y Scalabrini Ortiz · 19:50 hs', 'Palermo mapea al punto de Corrientes');
assert.strictEqual(api.lookup('PPMUMN').menor, true, 'Menor 18 pide celular del adulto');
assert.strictEqual(api.lookup('KXZKUQ').menor, false, 'Adulto no pide celular del adulto');
assert.strictEqual(api.lookup('WQWDEE').menor, true, 'sin rango de edad cargado, se pide igual');
assert.strictEqual(api.lookup('XUBPGE').apellido, '', 'el guion largo no es apellido');
assert.strictEqual(api.lookup('ZZZZZZ').error, 'invalido');
assert.strictEqual(api.lookup('  kxz kuq ').ok, true, 'normaliza espacios y minúsculas');
assert.strictEqual(api.lookup('KXZ').error, 'invalido', 'código corto');

// confirmar: caso feliz
e = entorno(base());
r = api.confirmar({
  code: 'KXZKUQ', nombre: 'Clara', apellido: 'Herrera', telefono: '11 2677-8578',
  restriccion: 'Ninguna', transfer: 'Sí', ida: 'Sí', vuelta: 'No', cancion: 'Súperestrella'
});
assert.strictEqual(r.ok, true, 'confirma');
let fila = e.invitados.datos[1];
const col = n => fila[CABECERA.indexOf(n)];
assert.strictEqual(col('Assistance_Confirmation'), 'Sí, ahí voy a estar');
assert.strictEqual(col('Name_Invite'), 'Clara');
assert.strictEqual(col('Phone'), '1126778578', 'el teléfono se guarda sin separadores');
assert.strictEqual(col('Adult_phone'), '', 'adulto no guarda celular de responsable');
assert.strictEqual(col('Transfer_use_outbound'), 'Sí');
assert.strictEqual(col('Transfer_use_inbound'), 'No');
assert.strictEqual(col('Song_preferece'), 'Súperestrella');
assert.ok(col('Timestamp') instanceof Date, 'queda marcado como usado');

// el código sigue entrando: lo que cambia es el estado de la fila
assert.strictEqual(api.lookup('KXZKUQ').estado, 'confirmado', 'quien ya contestó entra como confirmado');
assert.strictEqual(api.lookup('PPMUMN').estado, 'nuevo', 'sin marca de tiempo, es nuevo');
assert.strictEqual(api.lookup('QNNTVJ').punto, '', 'sin sitio cargado, no hay punto que ofrecer');

// lo guardado vuelve con los nombres del formulario
let prev = api.lookup('KXZKUQ').respuestas;
assert.deepStrictEqual(
  [prev.nombre, prev.telefono, prev.restriccion, prev.transfer, prev.ida, prev.vuelta, prev.cancion],
  ['Clara', '1126778578', 'Ninguna', 'Sí', 'Sí', 'No', 'Súperestrella'], 'la precarga devuelve lo que se guardó');
assert.strictEqual(api.lookup('PPMUMN').respuestas.nombre, '', 'una fila sin respuestas no precarga nada');

// al adulto ni se le pide: si igual llega, no se guarda
e = entorno(base());
api.confirmar({ code: 'KXZKUQ', telefono: '1126778578', adultoTel: '1199999999', restriccion: 'Ninguna', transfer: 'No' });
assert.strictEqual(e.invitados.datos[1][CABECERA.indexOf('Adult_phone')], '', 'el adulto nunca guarda celular de responsable');

// menor sin celular del adulto
e = entorno(base());
r = api.confirmar({ code: 'PPMUMN', telefono: '1126778578', restriccion: 'Ninguna', transfer: 'No' });
assert.deepStrictEqual([r.error, r.campo], ['incompleto', 'adultoTel'], 'menor sin celular del adulto');
r = api.confirmar({ code: 'PPMUMN', telefono: '1126778578', adultoTel: '1134558976', restriccion: 'Ninguna', transfer: 'No' });
assert.strictEqual(r.ok, true, 'menor con celular del adulto');
assert.strictEqual(e.invitados.datos[2][CABECERA.indexOf('Adult_phone')], '1134558976');

// el navegador no puede cambiar el punto asignado
e = entorno(base());
api.confirmar({ code: 'KXZKUQ', telefono: '1126778578', restriccion: 'Ninguna', transfer: 'Sí', ida: 'Sí', vuelta: 'Sí', punto: 'Palermo · a las 3 am' });
assert.strictEqual(e.invitados.datos[1][CABECERA.indexOf('Transfer_Site')], 'Recoleta', 'el punto de la planilla manda');

// sin punto cargado no hay transfer, aunque el navegador insista
e = entorno(base());
r = api.confirmar({ code: 'QNNTVJ', nombre: 'Carmen', apellido: 'Martearena', telefono: '1126778578', restriccion: 'Ninguna', transfer: 'Sí', ida: 'Sí', vuelta: 'No' });
assert.strictEqual(r.ok, true, 'confirma sin transfer');
assert.strictEqual(e.invitados.datos[3][CABECERA.indexOf('Transfer_Site')], '', 'no completa el sitio faltante');
['Transfer_use_outbound', 'Transfer_use_inbound'].forEach(c =>
  assert.strictEqual(e.invitados.datos[3][CABECERA.indexOf(c)], '', c + ' queda vacío si no se le ofrece transfer'));
assert.strictEqual(e.invitados.datos[3][CABECERA.indexOf('Surname_Invite')], 'Martearena', 'guarda el apellido escrito a mano');
assert.strictEqual(e.invitados.datos[3][CABECERA.indexOf('Surname_DB')], '', 'no toca el dato original');

// sin respuesta de transfer no se inventa una: las celdas quedan vacías, no en "No"
e = entorno(base());
assert.strictEqual(api.confirmar({ code: 'KXZKUQ', telefono: '1126778578', restriccion: 'Ninguna' }).ok, true);
['Transfer_use_outbound', 'Transfer_use_inbound'].forEach(c =>
  assert.strictEqual(e.invitados.datos[1][CABECERA.indexOf(c)], '', c + ' vacío si no contestó'));

// rechazar el transfer sí deja "No" en las dos
e = entorno(base());
api.confirmar({ code: 'KXZKUQ', telefono: '1126778578', restriccion: 'Ninguna', transfer: 'No' });
assert.strictEqual(e.invitados.datos[1][CABECERA.indexOf('Transfer_use_outbound')], 'No');

// códigos duplicados: corta en vez de elegir uno
e = entorno(base().concat([filaVacia(200, 'KXZKUQ', 'Otra', 'Persona', 'Adulto', 'Recoleta')]));
assert.strictEqual(api.lookup('KXZKUQ').error, 'duplicado');

// falta una columna: lo dice en vez de escribir en la que no es
const recortada = [CABECERA.filter(c => c !== 'Phone')].concat([filaVacia(3, 'KXZKUQ', 'Clara', 'Herrera', 'Adulto', 'Recoleta').slice(0, CABECERA.length - 1)]);
global.SpreadsheetApp = { getActive: () => ({ getSheets: () => [hojaFalsa('x', recortada)], getSheetByName: () => null, insertSheet: n => hojaFalsa(n, []) }) };
r = api.lookup('KXZKUQ');
assert.strictEqual(r.error, 'config');
assert.deepStrictEqual(r.falta, ['Phone']);

// rechazar el transfer deja "No" en las dos, distinto del vacío de quien no lo tiene
e = entorno(base());
api.confirmar({ code: 'KXZKUQ', telefono: '1126778578', restriccion: 'Ninguna', transfer: 'No' });
assert.strictEqual(e.invitados.datos[1][CABECERA.indexOf('Transfer_use_outbound')], 'No');
assert.strictEqual(e.invitados.datos[1][CABECERA.indexOf('Transfer_use_inbound')], 'No');

// el log registra todo
e = entorno(base());
api.confirmar({ code: 'ZZZZZZ' });
api.confirmar({ code: 'KXZKUQ', telefono: '1126778578', restriccion: 'Ninguna', transfer: 'No' });
const log = e.libro.getSheetByName('Log').datos;
assert.strictEqual(log.length, 3, 'encabezado + dos intentos');
assert.deepStrictEqual(log[0], ['Fecha', 'Codigo', 'Accion', 'Antes', 'Despues']);
assert.strictEqual(log[1][2], 'rechazo:invalido');
assert.strictEqual(log[2][2], 'alta');
assert.strictEqual(JSON.parse(log[2][3]).Name_Invite, '', 'antes de un alta la fila está vacía');
assert.strictEqual(JSON.parse(log[2][4]).Name_Invite, 'Clara', 'el log guarda lo que se escribió');

// El formato lo valida la página. Acá solo el piso: nada vacío, nada gigante.
e = entorno(base());
const largo = n => 'a'.repeat(n);
const enviar = extra => api.confirmar(Object.assign(
  { code: 'KXZKUQ', telefono: '1126778578', restriccion: 'Ninguna', transfer: 'No' }, extra || {}));

assert.strictEqual(enviar({ telefono: '' }).campo, 'telefono', 'sin teléfono no se guarda');
assert.strictEqual(enviar({ restriccion: '' }).campo, 'restriccion', 'sin restricción no se guarda');
assert.strictEqual(enviar({ cancion: largo(501) }).campo, 'largo', 'no se puede reventar una celda');
assert.strictEqual(enviar({ nombre: largo(501) }).campo, 'largo', 'ni con el nombre');
assert.strictEqual(enviar({ cancion: largo(500) }).ok, true, 'quinientos justos entran');

// el menor sigue necesitando el celular del adulto: eso lo decide la planilla, no el form
e = entorno(base());
assert.deepStrictEqual([enviar({ code: 'PPMUMN' }).error, enviar({ code: 'PPMUMN' }).campo],
  ['incompleto', 'adultoTel'], 'menor sin celular del adulto');

// el teléfono se guarda normalizado, con el + si venía del exterior
e = entorno(base());
assert.strictEqual(enviar({ telefono: '11 2677-8578' }).ok, true, 'guiones y espacios entran');
assert.strictEqual(e.invitados.datos[1][CABECERA.indexOf('Phone')], '1126778578', 'se guarda sin separadores');

e = entorno(base());
assert.strictEqual(enviar({ telefono: '+55 11 91234-5678' }).ok, true, 'número del exterior');
assert.strictEqual(e.invitados.datos[1][CABECERA.indexOf('Phone')], '+5511912345678', 'guarda el + y solo los dígitos');


// ---- Modificar, cancelar, volver a confirmar ----

const alta = extra => api.confirmar(Object.assign(
  { code: 'KXZKUQ', nombre: 'Clara', apellido: 'Herrera', telefono: '1126778578',
    restriccion: 'Ninguna', transfer: 'Sí', ida: 'Sí', vuelta: 'Sí', cancion: 'Súperestrella' }, extra || {}));
const dato = (fila, nombre) => e.invitados.datos[fila][CABECERA.indexOf(nombre)];

// una modificación pisa lo anterior
e = entorno(base());
alta();
r = alta({ telefono: '1155667788', restriccion: 'Kosher', transfer: 'No', cancion: '' });
assert.strictEqual(r.ok, true, 'la segunda vez también entra');
assert.strictEqual(r.estado, 'confirmado');
assert.strictEqual(dato(1, 'Phone'), '1155667788', 'el teléfono nuevo pisa al viejo');
assert.strictEqual(dato(1, 'Dietary_restrictions'), 'Kosher');
assert.strictEqual(dato(1, 'Transfer_use_outbound'), 'No', 'el transfer rechazado pisa al aceptado');
assert.strictEqual(dato(1, 'Song_preferece'), '', 'un campo vaciado se vacía en la planilla');
assert.strictEqual(e.libro.getSheetByName('Log').datos[2][2], 'modificacion');

// cancelar borra las respuestas y deja el motivo
e = entorno(base());
alta();
r = api.cancelar({ code: 'KXZKUQ' });
assert.strictEqual(r.ok, true, 'cancela');
assert.strictEqual(r.estado, 'cancelado');
assert.strictEqual(dato(1, 'Assistance_Confirmation'), 'No voy a poder ir');
['Name_Invite', 'Surname_Invite', 'Phone', 'Adult_phone', 'Transfer_use_outbound',
 'Transfer_use_inbound', 'Dietary_restrictions', 'Dietary_restrictions_other', 'Song_preferece']
  .forEach(c => assert.strictEqual(dato(1, c), '', c + ' se borra al cancelar'));
assert.ok(dato(1, 'Timestamp') instanceof Date, 'la cancelación también deja marca de tiempo');
assert.strictEqual(dato(1, 'Name_DB'), 'Clara', 'el dato original no se toca');

// lo cancelado sigue entrando, y lo que había puesto se rescata del log
assert.strictEqual(api.lookup('KXZKUQ').estado, 'cancelado');
assert.strictEqual(api.lookup('KXZKUQ').respuestas.telefono, '', 'después de cancelar no queda nada que precargar');
const cancelacion = e.libro.getSheetByName('Log').datos[2];
assert.strictEqual(cancelacion[2], 'cancelacion');
assert.strictEqual(JSON.parse(cancelacion[3]).Phone, '1126778578', 'el log conserva lo que se borró');
assert.strictEqual(JSON.parse(cancelacion[4]).Assistance_Confirmation, 'No voy a poder ir');

// cancelar dos veces no vuelve a escribir ni ensucia el log
r = api.cancelar({ code: 'KXZKUQ' });
assert.deepStrictEqual([r.ok, r.estado], [true, 'cancelado'], 'cancelar de nuevo no falla');
assert.strictEqual(e.libro.getSheetByName('Log').datos.length, 3, 'no registra una cancelación que no cambió nada');

// y después de cancelar se puede volver a confirmar
r = alta({ telefono: '1144332211' });
assert.strictEqual(r.ok, true, 'vuelve a confirmar');
assert.strictEqual(dato(1, 'Assistance_Confirmation'), 'Sí, ahí voy a estar');
assert.strictEqual(dato(1, 'Phone'), '1144332211');
assert.strictEqual(e.libro.getSheetByName('Log').datos[3][2], 'reconfirmacion');

// cancelar un código que no existe no escribe nada
e = entorno(base());
assert.strictEqual(api.cancelar({ code: 'ZZZZZZ' }).error, 'invalido');
assert.strictEqual(e.libro.getSheetByName('Log').datos[1][2], 'rechazo:invalido');

// cancelar antes de contestar deja la fila como quien dijo que no
e = entorno(base());
assert.strictEqual(api.cancelar({ code: 'KXZKUQ' }).estado, 'cancelado');
assert.strictEqual(dato(1, 'Assistance_Confirmation'), 'No voy a poder ir');

console.log('Codigo.gs: todas las pruebas pasaron');
