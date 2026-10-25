/* ============================================================
   AVENTURApp — backend (Google Apps Script)
   Se pega entero en el editor de Apps Script del Sheet.

   Primera vez: ejecutar setup() desde el editor (crea pestañas,
   guarda el ID del Sheet y siembra la aventura de ejemplo).
   Luego: Implementar › Nueva implementación › Aplicación web
          Ejecutar como: yo · Acceso: cualquier usuario.
   Cada cambio de código: Gestionar implementaciones › ✏️ › Versión nueva.
   ============================================================ */
const VERSION = '0.3.0';

const HEADERS = {
  Usuarios:  ['username', 'nombre', 'avatar', 'color', 'activo'],
  Aventuras: ['id', 'titulo', 'epoca', 'intro', 'color', 'emoji', 'dibujo', 'orden', 'activa',
              'kit', 'fondo', 'prota', 'elementos', 'imagen'],
  Escenas:   ['aventura_id', 'num', 'tipo', 'titulo', 'texto', 'dato',
              'opcion_a', 'destino_a', 'opcion_b', 'destino_b', 'dibujo', 'emoji', 'sonido',
              'fondo', 'prota', 'elementos', 'imagen'],
  Partidas:  ['id', 'ts', 'aventura_id', 'recorrido', 'final', 'jugadora'],
  Config:    ['clave', 'valor', 'nota']
};

/* getActiveSpreadsheet() devuelve null cuando corre como Web App:
   setup() guarda el ID y a partir de ahí se abre por ID. */
function _ss() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function _cell(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  return v;
}
/** Lee una pestaña como array de objetos, por nombre de cabecera. Si no existe → []. */
function _readSheet(name) {
  const sh = _ss().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const head = vals[0].map(h => String(h).trim());
  return vals.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => { const o = {}; head.forEach((h, i) => { if (h) o[h] = _cell(r[i]); }); return o; });
}
/** Añade una fila respetando el orden real de las cabeceras. */
function _append(name, obj) {
  const sh = _ss().getSheetByName(name);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  sh.appendRow(head.map(h => (obj[h] === undefined || obj[h] === null) ? '' : obj[h]));
}

/* ---------- puerta de entrada ---------- */
function doGet(e) { return _despachar((e && e.parameter) || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return _despachar(p);
}
const HANDLERS = {
  ping:           () => _json({ ok: true, data: { version: VERSION } }),
  getBootstrap:   handleGetBootstrap,
  savePartida:    handleSavePartida,
  borrarPartidas: handleBorrarPartidas
};
function _despachar(p) {
  const h = HANDLERS[p.action];
  if (!h) return _json({ ok: false, error: 'Acción desconocida: ' + p.action });
  try { return h(p); }
  catch (err) { return _json({ ok: false, error: String(err && err.message || err) }); }
}

/* ---------- handlers ---------- */
function handleGetBootstrap() {
  return _json({ ok: true, data: {
    version:   VERSION,
    config:    _readSheet('Config'),
    usuarios:  _readSheet('Usuarios'),
    aventuras: _readSheet('Aventuras'),
    escenas:   _readSheet('Escenas'),
    partidas:  _readSheet('Partidas')
  }});
}
/** Idempotente por id: si la cola reenvía una partida, no se duplica. */
function handleSavePartida(p) {
  const d = typeof p.payload === 'string' ? JSON.parse(p.payload) : (p.payload || {});
  if (!d.id || !d.aventura_id || !d.final) return _json({ ok: false, error: 'Faltan datos de la partida.' });
  const ya = _readSheet('Partidas').some(r => String(r.id) === String(d.id));
  if (!ya) _append('Partidas', {
    id: d.id, ts: d.ts || new Date().toISOString(), aventura_id: d.aventura_id,
    recorrido: "'" + String(d.recorrido || ''), /* el apóstrofo evita que Sheets lo lea como fecha */
    final: Number(d.final), jugadora: d.jugadora || ''
  });
  return _json({ ok: true, data: { id: d.id } });
}
/** Destructiva: solo actúa si llega confirmar=si. Borra las partidas de UNA jugadora. */
function handleBorrarPartidas(p) {
  if (String(p.confirmar) !== 'si') return _json({ ok: false, error: 'Falta confirmar.' });
  const quien = String(p.jugadora || '').trim().toLowerCase();
  if (!quien) return _json({ ok: false, error: 'Falta la jugadora.' });
  const sh = _ss().getSheetByName('Partidas');
  if (!sh || sh.getLastRow() < 2) return _json({ ok: true, data: { borradas: 0 } });
  const vals = sh.getDataRange().getValues(), col = vals[0].map(String).indexOf('jugadora');
  let n = 0;
  for (let r = vals.length - 1; r >= 1; r--) {
    if (String(vals[r][col]).trim().toLowerCase() === quien) { sh.deleteRow(r + 1); n++; }
  }
  return _json({ ok: true, data: { borradas: n } });
}

/* ============================================================
   MANTENIMIENTO (se ejecutan a mano desde el editor)
   ============================================================ */

/** Crea pestañas y columnas que falten, guarda el ID del Sheet y siembra la aventura
    de ejemplo si Aventuras está vacía. Idempotente: se puede correr las veces que haga falta. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const cols = HEADERS[name];
    const actual = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    const faltan = cols.filter(c => actual.indexOf(c) < 0);
    if (faltan.length) sh.getRange(1, actual.filter(String).length + 1, 1, faltan.length).setValues([faltan]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
  });
  if (!_readSheet('Usuarios').length) {
    _append('Usuarios', { username: 'gina', nombre: 'Gina', avatar: 'kit/paleolitico/gina/contenta.webp', color: '#F6C445', activo: true });
  }
  if (!_readSheet('Config').length) {
    _append('Config', { clave: 'usuario_base', valor: 'gina', nota: 'Jugadora de las imágenes por defecto. Otras usan img/escenas/<usuario>/ si existe.' });
  }
  if (!_readSheet('Aventuras').length) _sembrar();
  Logger.log('✅ setup listo. Ahora: Implementar › Nueva implementación › Aplicación web.');
}

/** Revisa que cada aventura se pueda jugar: destinos que existen, decisiones con 2 opciones,
    finales alcanzables. Solo informa, no cambia nada. */
function diagnosticar() {
  const esc = _readSheet('Escenas');
  _readSheet('Aventuras').forEach(a => {
    const E = {};
    esc.filter(e => String(e.aventura_id) === String(a.id)).forEach(e => { E[Number(e.num)] = e; });
    const problemas = [];
    Object.keys(E).forEach(n => {
      const e = E[n], tipo = String(e.tipo || 'decision').toLowerCase();
      if (tipo === 'decision') {
        ['a', 'b'].forEach(k => {
          if (!e['opcion_' + k]) problemas.push('Escena ' + n + ': falta opcion_' + k);
          else if (!E[Number(e['destino_' + k])]) problemas.push('Escena ' + n + ': destino_' + k + ' (' + e['destino_' + k] + ') no existe');
        });
      } else if (tipo !== 'fin' && tipo !== 'final') problemas.push('Escena ' + n + ': tipo "' + e.tipo + '" no vale (decision / fin / final)');
    });
    /* alcanzables desde la 1 */
    const vis = {}, q = [1];
    while (q.length) { const n = q.shift(); if (vis[n] || !E[n]) continue; vis[n] = true; q.push(Number(E[n].destino_a), Number(E[n].destino_b)); }
    Object.keys(E).forEach(n => { if (!vis[n]) problemas.push('Escena ' + n + ': no se puede llegar desde la 1'); });
    if (!Object.keys(E).some(n => String(E[n].tipo).toLowerCase() === 'final')) problemas.push('No hay ningún final feliz (tipo "final")');
    Logger.log((problemas.length ? '⚠️ ' : '✅ ') + a.id + ' — ' + Object.keys(E).length + ' escenas' +
      (problemas.length ? '\n   ' + problemas.join('\n   ') : ', todo bien'));
  });
}

/** Vuelve a escribir el Paleolítico con la versión de este código (textos, composición e
    imágenes). Solo informa; resembrarPaleoliticoAplicar() borra sus filas y las reescribe. */
function resembrarPaleolitico(aplicar) {
  const ss = _ss();
  const n = _readSheet('Escenas').filter(e => String(e.aventura_id) === 'paleolitico').length;
  if (!aplicar) { Logger.log('Hay ' + n + ' escenas de "paleolitico". Ejecuta resembrarPaleoliticoAplicar() para reescribirlas.'); return; }
  ['Aventuras', 'Escenas'].forEach(tab => {
    const sh = ss.getSheetByName(tab), vals = sh.getDataRange().getValues();
    const col = vals[0].map(String).indexOf(tab === 'Aventuras' ? 'id' : 'aventura_id');
    for (let r = vals.length - 1; r >= 1; r--) if (/^(hielo|paleolitico)$/.test(String(vals[r][col]))) sh.deleteRow(r + 1);
  });
  _sembrar();
  Logger.log('✅ Paleolítico reescrito.');
}
function resembrarPaleoliticoAplicar() { resembrarPaleolitico(true); }

function _sembrar() {
  _append('Aventuras', { id: 'paleolitico', titulo: 'Paleolítico', epoca: 'Hace 15.000 años · Edad de Hielo',
    intro: 'Vives con tu clan de cazadores-recolectores. Hace frío, hay mamuts… y todos tienen hambre. ¿Sobrevivirás?',
    color: '#9CCBE8', emoji: '🦣', dibujo: 'portada_hielo', orden: 1, activa: true,
    kit: 'paleolitico', fondo: 'montanas', prota: 'entusiasmada', elementos: 'mamut', imagen: 'paleolitico_portada.webp' });
  const filas = [
    [1,'decision','¡A por comida!','Hace 15.000 años, en plena Edad de Hielo. Tu clan tiene hambre y hoy tú vas a ayudar. ¿Qué te llevas?','Hacía mucho más frío que hoy y el norte de Europa estaba cubierto de hielo.','Lanza y propulsor',2,'Bolsa de piel para recoger frutos',3,'campamento','🏕️','','campamento','normal','bolsa'],
    [2,'decision','¡Renos!','Una manada de renos cruza el valle. ¡Comida para muchos días!','Los renos eran una de sus presas favoritas: daban carne, piel, huesos y astas.','Acercarte despacio, con el viento de cara',5,'Correr gritando «¡A POR ELLOS!»',4,'renos','🦌','','valle','entusiasmada','reno,reno'],
    [3,'decision','¿Qué bayas?','Hay unas bayas rojas superbrillantes y otras oscuras: las que siempre coge la abuela.','Además de cazar, recolectaban bayas, raíces, frutos secos y setas.','Las rojas, ¡qué bonitas!',6,'Las de la abuela',7,'bayas','🫐','','bosque','pensando','abuela,bayas'],
    [4,'fin','¡Estampida!','Los renos se asustan y salen corriendo… ¡por encima de ti! Te quedas plana como una tortita.','Cazaban en silencio y con el viento de cara, para que los animales no los olieran.','','','','','estampida','🦌','¡PLAF!','valle','mareada','reno,huellas'],
    [5,'decision','¡Un mamut!','Un mamut lanudo, grande como un elefante, te tapa el camino. Come hierba tan tranquilo.','El mamut lanudo tenía pelo largo, orejas pequeñas y una capa de grasa para aguantar el frío.','Pasar por debajo: ¡cabes!',8,'Rodearlo despacito',9,'mamut','🦣','','nieve','sorpresa','mamut'],
    [6,'fin','Bayas traicioneras','¡Eran venenosas! Pasas tres días con dolor de barriga… escondida detrás de un arbusto.','Los mayores enseñaban a los niños qué plantas se podían comer y cuáles no.','','','','','barriga','🤢','¡GRRRUÑ!','bosque','asco','bayas'],
    [7,'decision','¡Tormenta!','Ya tienes la bolsa llena, pero empieza a nevar muchísimo y se hace de noche.','Con tormenta buscaban refugio en cuevas y abrigos de roca.','Meterte en una cueva',10,'Seguir andando a ciegas',11,'tormenta','🌨️','','nieve','miedo','copo,cueva'],
    [8,'fin','¡Pisotón!','El mamut ni te ve. Da un pasito… y te quedas plana como una alfombra.','Un mamut lanudo pesaba más que cuatro coches juntos. ¡Mejor mirarlo de lejos!','','','','','pisoton','🦣','¡PLOF!','valle','mareada','mamut,estrella'],
    [9,'decision','¡Zas!','Papá lanza con el propulsor y caza un reno. ¡Bravo! ¿Cómo lo lleváis al campamento?','El propulsor es un palo con gancho que hace volar la lanza más lejos y más rápido.','Despiezarlo y cargarlo entre todos',12,'Dejarlo solo e ir a por ayuda',13,'zas','🎯','','valle','celebrando','papa,reno'],
    [10,'decision','¡Qué oscuro!','Dentro de la cueva no se ve nada. Y al fondo se oye un ruido muy raro…','Hacían fuego golpeando sílex contra pirita: ¡saltan chispas!','Hacer fuego con sílex y pirita',14,'Ir al fondo a ver qué es',15,'oscuro','🕳️','','cueva_dentro','miedo','interrogacion'],
    [11,'fin','Hielo fino','Pisas hielo fino y… ¡al agua helada! Sales tiesa como un polo de hielo.','Para no congelarse llevaban ropa de piel cosida con agujas de hueso.','','','','','hielo','🧊','¡CRAAAC!','rio','llorando','nieve,gotas'],
    [12,'decision','El río helado','Lleváis carne, piel, huesos y tendones: ¡todo se aprovecha! Ahora hay que cruzar el río.','Con los tendones hacían hilo, y con los huesos, agujas y herramientas.','Por el hielo: ¡es más corto!',11,'Saltando por las piedras',16,'rio','🏞️','','rio','contenta','carne,piedras'],
    [13,'fin','¡Un león!','Un león de las cavernas huele la comida. Corres tanto que llegas a casa… ¡sin nada!','Los leones de las cavernas casi no tenían melena. ¡Así los pintaron en la cueva de Chauvet (Francia)!','','','','','leon','🦁','¡GRRR!','bosque','corriendo','leon,huellas'],
    [14,'decision','¡Pinturas!','Con la luz del fuego ves bisontes pintados en la pared. ¡Qué pasada! Al amanecer, toca volver.','Iluminaban las cuevas con lámparas de piedra que quemaban grasa animal.','Seguir unas huellas de gato gigante',13,'Volver por el camino de siempre',16,'pinturas','🎨','','cueva_dentro','asombrada','antorcha'],
    [15,'fin','La osa dormilona','Era una osa durmiendo. Te sientas encima sin querer… y se despierta de MUY mal humor.','Los osos pardos pasan el invierno durmiendo dentro de las cuevas.','','','','','osa','🐻','¡UY!','cueva_dentro','asustada','oso,zzz'],
    [16,'final','¡Lo lograste!','Toda la familia cena junto al fuego, el abuelo toca una flauta de hueso y tú pintas tu mano en la cueva. ¡Eres una aventurera de la Edad de Hielo!','Pintaban manos soplando ocre alrededor, como en la cueva de El Castillo (Cantabria).','','','','','fiesta','🔥','¡HURRA!','campamento','celebrando','fuego,mama,papa']
  ];
  filas.forEach(r => _append('Escenas', { aventura_id: 'paleolitico', num: r[0], tipo: r[1], titulo: r[2], texto: r[3], dato: r[4],
    opcion_a: r[5], destino_a: r[6], opcion_b: r[7], destino_b: r[8], dibujo: r[9], emoji: r[10], sonido: r[11],
    fondo: r[12], prota: r[13], elementos: r[14], imagen: 'paleolitico_' + ('0' + r[0]).slice(-2) + '.webp' }));
}
