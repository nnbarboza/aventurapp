/* ============================================================
   AVENTURApp — backend (Google Apps Script)
   Se pega entero en el editor de Apps Script del Sheet.

   Primera vez: ejecutar setup() desde el editor (crea pestañas,
   guarda el ID del Sheet y siembra la aventura de ejemplo).
   Luego: Implementar › Nueva implementación › Aplicación web
          Ejecutar como: yo · Acceso: cualquier usuario.
   Cada cambio de código: Gestionar implementaciones › ✏️ › Versión nueva.
   ============================================================ */
const VERSION = '0.14.0';

const HEADERS = {
  Usuarios:  ['username', 'nombre', 'avatar', 'color', 'activo'],
  Aventuras: ['id', 'titulo', 'epoca', 'intro', 'color', 'emoji', 'dibujo', 'orden', 'activa',
              'kit', 'fondo', 'prota', 'elementos', 'imagen', 'hace', 'categoria', 'quiz_n', 'mochila', 'mochila_emoji'],
  Escenas:   ['aventura_id', 'num', 'tipo', 'titulo', 'texto', 'dato',
              'opcion_a', 'destino_a', 'opcion_b', 'destino_b', 'dibujo', 'emoji', 'sonido',
              'fondo', 'prota', 'elementos', 'imagen',
              'obtiene', 'opcion_c', 'destino_c', 'requiere_a', 'sino_a', 'requiere_b', 'sino_b', 'requiere_c',
              'obtiene_a', 'obtiene_b', 'obtiene_c', 'gasta_a', 'gasta_b', 'gasta_c'],
  Objetos:   ['aventura_id', 'id', 'nombre', 'emoji', 'imagen', 'descripcion', 'dato'],
  Partidas:  ['id', 'ts', 'aventura_id', 'recorrido', 'final', 'jugadora', 'objetos'],
  Quiz:      ['aventura_id', 'escena', 'pregunta', 'correcta', 'incorrecta_1', 'incorrecta_2', 'nivel'],
  Quizzes:   ['id', 'ts', 'aventura_id', 'aciertos', 'total', 'jugadora'],
  Invitados: ['username', 'nombre', 'avatar', 'creado', 'ultimo'],
  Secretos:  ['jugadora', 'aventura_id', 'num', 'ts'],
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
  saveQuiz:       handleSaveQuiz,
  saveInvitado:   handleSaveInvitado,
  saveSecreto:    handleSaveSecreto,
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
    partidas:  _readSheet('Partidas'),
    quiz:      _readSheet('Quiz'),
    quizzes:   _readSheet('Quizzes'),
    objetos:   _readSheet('Objetos'),
    secretos:  _readSheet('Secretos')
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
    final: Number(d.final), jugadora: d.jugadora || '', objetos: d.objetos || ''
  });
  return _json({ ok: true, data: { id: d.id } });
}
/** Resultado de un quiz. Idempotente por id. */
function handleSaveQuiz(p) {
  const d = typeof p.payload === 'string' ? JSON.parse(p.payload) : (p.payload || {});
  if (!d.id || !d.aventura_id) return _json({ ok: false, error: 'Faltan datos del quiz.' });
  if (!_readSheet('Quizzes').some(r => String(r.id) === String(d.id))) _append('Quizzes', {
    id: d.id, ts: d.ts || new Date().toISOString(), aventura_id: d.aventura_id,
    aciertos: Number(d.aciertos) || 0, total: Number(d.total) || 0, jugadora: d.jugadora || ''
  });
  return _json({ ok: true, data: { id: d.id } });
}

/** Un secreto de Chewie leído (jugadora + aventura + escena). Idempotente. */
function handleSaveSecreto(p) {
  const d = typeof p.payload === 'string' ? JSON.parse(p.payload) : (p.payload || {});
  if (!d.jugadora || !d.aventura_id || !d.num) return _json({ ok: false, error: 'Faltan datos del secreto.' });
  const ya = _readSheet('Secretos').some(r => String(r.jugadora) === String(d.jugadora) && String(r.aventura_id) === String(d.aventura_id) && Number(r.num) === Number(d.num));
  if (!ya) _append('Secretos', { jugadora: d.jugadora, aventura_id: d.aventura_id, num: Number(d.num), ts: d.ts || new Date().toISOString() });
  return _json({ ok: true, data: {} });
}

/** Apunta (o actualiza) un invitado: nombre, avatar, cuándo se creó y la última vez que jugó. */
function handleSaveInvitado(p) {
  const d = typeof p.payload === 'string' ? JSON.parse(p.payload) : (p.payload || {});
  const u = String(d.username || '').trim();
  if (!u) return _json({ ok: false, error: 'Falta el invitado.' });
  const ts = d.ts || new Date().toISOString();
  const sh = _ss().getSheetByName('Invitados'), vals = sh.getDataRange().getValues(), head = vals[0].map(String);
  const iU = head.indexOf('username');
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][iU]) === u) {
      sh.getRange(r + 1, head.indexOf('nombre') + 1).setValue(d.nombre || vals[r][head.indexOf('nombre')]);
      sh.getRange(r + 1, head.indexOf('avatar') + 1).setValue(d.avatar || vals[r][head.indexOf('avatar')]);
      sh.getRange(r + 1, head.indexOf('ultimo') + 1).setValue(ts);
      return _json({ ok: true, data: { username: u } });
    }
  }
  _append('Invitados', { username: u, nombre: d.nombre || '', avatar: d.avatar || '', creado: ts, ultimo: ts });
  return _json({ ok: true, data: { username: u } });
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
  _completarV06();
  _completarV09();
  _completarV10();
  _completarV11();
  _completarV12();
  _completarV13();
  Logger.log('✅ setup listo. Ahora: Implementar › Nueva implementación › Aplicación web.');
}

/* ---------- v0.6: Neolítico, quiz, línea del tiempo ---------- */
const NEO_AVENTURA = {"id": "neolitico", "titulo": "Neolítico", "epoca": "Hace 7.300 años · La Draga (Girona)", "intro": "Tu familia vive en La Draga, un poblado junto al lago de Banyoles. Ya no sois cazadores: ¡ahora sembráis, criáis cabras y hacéis cerámica! ¿Llegarás a la gran cosecha?", "color": "#B7D99B", "emoji": "🌾", "dibujo": "", "orden": 2, "activa": true, "kit": "neolitico", "hace": 7300, "imagen": "neolitico_portada.webp"};
const METALES_PRONTO = {"id": "metales", "titulo": "Edad de los Metales", "epoca": "Cobre y bronce · hace 5.000 años", "intro": "Pronto: hornos, metales y los primeros poblados fortificados.", "color": "#E8C48A", "emoji": "⚒️", "orden": 3, "activa": "pronto", "hace": 5000};
const NEO_ESCENAS = [
  [1, "decision", "Primavera en La Draga", "Despiertas en tu casa de madera, levantada sobre pilotes a la orilla del lago. Es primavera y en el poblado hay muchísimo trabajo. ¿Con quién vas?", "La Draga, junto al lago de Banyoles (Girona), es uno de los primeros poblados de agricultores de la Península. Tiene unos 7.300 años.", "Con papá, a sembrar el campo", 2, "Con el abuelo, a cuidar las cabras", 3, "🏡", ""],
  [2, "decision", "¡A sembrar!", "Papá hace agujeros en la tierra con un palo cavador. Te da una bolsa de semillas de trigo… pero tú llevas otra con bellotas que encontraste ayer.", "Sembraban trigo y cebada con semillas guardadas de la cosecha anterior.", "Sembrar el trigo de papá", 4, "Sembrar tus bellotas: ¡crecerán más rápido!", 5, "🌱", ""],
  [3, "decision", "Las cabras", "El abuelo te deja vigilar las cabras. De repente, la más traviesa sale corriendo hacia el lago.", "La cabra y la oveja llegaron con los primeros agricultores: se domesticaron en Oriente Próximo.", "Correr detrás gritando «¡VUELVE!»", 6, "Llamarla bajito y ofrecerle unas hojas", 7, "🐐", ""],
  [4, "decision", "Gorriones hambrientos", "Una bandada de gorriones mira tus semillas desde un árbol. ¡Se están relamiendo!", "Trabajaban la tierra con palos cavadores de madera: todavía no existía el arado.", "Dejarlas encima: así les da el sol", 9, "Taparlas bien con tierra", 8, "🐦", ""],
  [5, "fin", "Un bosque de robles", "En otoño, en vez de trigo… ¡te salen robles bebés! Tardarán cien años en crecer, y no dan pan.", "En la prehistoria también se recogían bellotas para comer… pero no se sembraban.", "", 0, "", 0, "🌳", "¡OPS!"],
  [6, "fin", "¡Excursión de cabras!", "Tus gritos asustan a TODAS las cabras, que se van de excursión por el bosque. Pasas la tarde buscándolas, con el abuelo resoplando detrás.", "En La Draga se han encontrado huesos de cabras, ovejas, vacas y cerdos.", "", 0, "", 0, "🐐", "¡BEEE!"],
  [7, "decision", "La cabra glotona", "La cabra vuelve a por las hojas y, detrás, aparece un perro del poblado moviendo la cola. ¡Qué ayuda!", "El perro fue el primer animal domesticado, miles de años antes que las cabras.", "Volver con el rebaño al poblado", 10, "Dejar al perro a cargo y echarte una siesta", 11, "🐕", ""],
  [8, "decision", "¡Sembrado!", "Papá está orgulloso: has sembrado como una experta. Antes de volver, te ofrece las semillas que le sobran.", "Los campos estaban cerca del poblado, para poder vigilarlos todos los días.", "Plantarlas todas en el mismo agujero", 12, "Guardarlas para otro día", 10, "👍", ""],
  [9, "fin", "Banquete de gorriones", "Los gorriones se dan el festín de su vida. Te dejan una pluma de regalo… y ni una sola semilla.", "Si se perdía la cosecha, en invierno no había comida: por eso cuidaban tanto los campos.", "", 0, "", 0, "🐦", "¡PÍO!"],
  [10, "decision", "Mediodía en el poblado", "Mamá cocina gachas de cereal en una olla de cerámica. ¡Qué rico! Después de comer, ¿qué haces?", "Molían el grano con piedras de moler para hacer gachas y tortas.", "Hacer cerámica con la abuela", 13, "Ir al bosque con papá a buscar madera", 14, "🥣", ""],
  [11, "fin", "Siesta desastrosa", "Cuando despiertas, las cabras se han comido el tejado… ¡de la casa de la vecina!", "Las casas eran de madera y los tejados, de ramas y plantas: ¡un bufé para las cabras!", "", 0, "", 0, "🏚️", "¡ÑAM!"],
  [12, "fin", "La torre de trigo", "Brotan cien plantas apretujadas en un solo agujero y ninguna crece. Papá se ríe tanto que se cae al barro.", "Cada semilla necesita su espacio, su agua y su luz para crecer.", "", 0, "", 0, "🌾", "¡CHOF!"],
  [13, "decision", "Barro y conchas", "La abuela modela una vasija de barro y te deja decorarla. ¿Con qué?", "Muchas vasijas se decoraban apretando el borde de una concha en el barro: se llama cerámica cardial.", "Con las pezuñas de una cabra", 16, "Con el borde de una concha, como la abuela", 15, "🏺", ""],
  [14, "decision", "El tejo", "En el bosque, papá busca madera de tejo para hacer arcos. Encontráis uno enorme.", "En La Draga se encontraron arcos de madera de tejo: ¡de los más antiguos de Europa!", "Trepar hasta las ramas más altas", 17, "Ayudar a papá a cortar una rama baja", 18, "🌲", ""],
  [15, "decision", "¡Al fuego!", "Tu vasija ha quedado preciosa. Ahora hay que cocerla en el fuego para que se endurezca.", "Las vasijas se secaban al sol antes de cocerlas en el fuego para endurecerlas.", "Esperar a que se seque al sol", 19, "Meterla al fuego ya, ¡que tengo prisa!", 20, "🔥", ""],
  [16, "fin", "La cabra alfarera", "Traes una cabra para que deje sus huellas… y pisotea TODAS las vasijas de la abuela. Clac, clac, clac.", "En La Draga se han encontrado muchísimos trozos de cerámica.", "", 0, "", 0, "🐐", "¡CLAC!"],
  [17, "fin", "Atascada en el tejo", "Subes tan alto que no sabes bajar. Papá tarda toda la tarde en rescatarte y se hace de noche sin madera.", "Las semillas del tejo son muy venenosas: ¡nunca se comen!", "", 0, "", 0, "🌲", "¡SOCORRO!"],
  [18, "decision", "El hacha pulida", "El hacha de papá es de piedra pulida, con mango de madera. ¡Zas! Cortáis una rama perfecta para un arco.", "Pulían la piedra para hacer hachas muy resistentes. Por eso esta época se llama Neolítico: «piedra nueva».", "Seguir unas huellas de jabalí", 22, "Volver al poblado con la rama", 21, "🪓", ""],
  [19, "decision", "Vasija perfecta", "Tu vasija sale del fuego dura y bonita. La abuela dice que es tan buena que servirá para guardar las semillas.", "Guardaban el grano en vasijas grandes o en hoyos en el suelo, llamados silos.", "Usarla de tambor en la fiesta del poblado", 24, "Guardarla en casa hasta la cosecha", 23, "🏺", ""],
  [20, "fin", "¡Pum!", "La vasija, todavía húmeda, explota en el fuego. ¡Llueven trocitos de barro sobre el abuelo!", "Si el barro aún tiene agua, al calentarse se convierte en vapor y lo rompe: por eso lo secaban antes.", "", 0, "", 0, "💥", "¡PUM!"],
  [21, "decision", "El arco", "En casa, papá talla el arco y tú lo pules con arena. Pasan los meses y llega el verano.", "Las flechas llevaban puntas de sílex muy afiladas.", "Ir con todos a segar el trigo", 23, "Ir de caza con papá y el arco nuevo", 25, "🏹", ""],
  [22, "fin", "¡Mamá jabalí!", "Encontráis al jabalí… y a sus seis jabatos. Mamá jabalí os persigue hasta el lago: ¡chapuzón con ropa!", "Los cerdos domésticos descienden del jabalí.", "", 0, "", 0, "🐗", "¡SPLASH!"],
  [23, "decision", "¡La cosecha!", "Es verano y el trigo está dorado. Todo el poblado sale a segar y a ti te dan una hoz de madera con dientes de sílex.", "Las hoces eran de madera, con trocitos de sílex afilado encajados como dientes.", "Segar a toda velocidad para ganar a tu compi", 27, "Segar con cuidado, como te enseñan", 26, "🌾", ""],
  [24, "fin", "Tambor roto", "La golpeas como un tambor en mitad del poblado: bum, bum, bum… ¡CRAC! Adiós, vasija.", "Los arqueólogos unen los trozos de cerámica como si fueran un puzle.", "", 0, "", 0, "🥁", "¡CRAC!"],
  [25, "decision", "La cacería", "Al amanecer, junto al lago, aparece un ciervo bebiendo agua. Papá te pasa el arco.", "Aunque ya eran agricultores, seguían cazando ciervos y jabalíes.", "Acercaros despacio, en silencio", 29, "Disparar desde lejísimos", 30, "🦌", ""],
  [26, "decision", "Guardar el grano", "¡Hay montañas de espigas! ¿Dónde guardáis el grano para el invierno?", "Antes de guardar el grano, lo separaban de la paja.", "En un montón en el suelo de casa", 31, "En vasijas bien tapadas", 28, "🫙", ""],
  [27, "fin", "La coleta del abuelo", "Vas tan rápido que segas trigo, hierba… ¡y la coleta del abuelo! Se queda con media melena.", "En la cosecha ayudaba todo el poblado: mayores, pequeños… y abuelos.", "", 0, "", 0, "✂️", "¡ZAS!"],
  [28, "decision", "Las mejores semillas", "Todo el grano está guardado. Los mayores preguntan qué hacer con las espigas más grandes y bonitas.", "Elegían las mejores espigas para sembrar: así, poco a poco, el trigo se hizo más grande.", "Guardarlas para sembrar en primavera", 32, "Hacer una fiesta y comérnoslas todas", 33, "🌾", ""],
  [29, "final", "¡Cena para el poblado!", "Papá caza el ciervo y todo el poblado lo celebra junto al lago. Esa noche, papá te regala tu primer arco, pequeñito. ¡Eres una cazadora de La Draga!", "Del ciervo se aprovechaba todo: carne, piel, tendones y astas para hacer herramientas.", "", 0, "", 0, "🏹", "¡HURRA!"],
  [30, "fin", "Flecha al lago", "Disparas desde lejísimos… la flecha cae al lago, el ciervo se va tan tranquilo y papá tiene que bucear para recuperarla.", "La madera de La Draga se ha conservado más de 7.000 años porque quedó bajo el agua y el barro.", "", 0, "", 0, "🎯", "¡PLOF!"],
  [31, "fin", "Fiesta de ratones", "Esa noche, los ratones del bosque celebran en tu casa la mejor fiesta de su vida.", "Guardar el grano seco y bien tapado era clave para sobrevivir al invierno.", "", 0, "", 0, "🐭", "¡ÑAM!"],
  [32, "final", "La gran cosecha", "Guardáis las mejores semillas en vasijas bien tapadas para sembrar en primavera. Esa noche, todo el poblado celebra con tortas de cereal recién hechas. ¡Eres una agricultora de La Draga!", "La agricultura nació hace más de 10.000 años en Oriente Próximo y tardó miles de años en llegar hasta aquí.", "", 0, "", 0, "🎉", "¡HURRA!"],
  [33, "fin", "Primavera sin semillas", "La fiesta es increíble… pero en primavera no queda ni una semilla para sembrar. ¡Ups! Toca ir a pedirlas a otro poblado.", "Los poblados intercambiaban cosas: semillas, piedras para herramientas, conchas…", "", 0, "", 0, "🙈", "¡UPS!"]
];
const QUIZ_SEMILLA = [
  ["paleolitico", 1, "¿Qué significa «Paleolítico»?", "Piedra antigua", "Piedra nueva", "Edad del hielo"],
  ["paleolitico", 2, "¿Qué les daban los renos?", "Carne, piel, huesos y astas", "Leche y queso", "Lana para tejer"],
  ["paleolitico", 5, "¿Qué ayudaba al mamut a aguantar el frío?", "Pelo largo y una capa de grasa", "Una manta de piel", "Dormir junto al fuego"],
  ["paleolitico", 9, "¿Para qué servía el propulsor?", "Para lanzar la lanza más lejos y más rápido", "Para encender fuego", "Para pescar"],
  ["paleolitico", 10, "¿Cómo hacían fuego?", "Golpeando sílex contra pirita", "Con cerillas", "Frotando dos huesos"],
  ["paleolitico", 11, "¿Cómo cosían la ropa de piel?", "Con agujas de hueso", "Con agujas de metal", "Con grapas"],
  ["paleolitico", 13, "¿Cómo eran los leones de las cavernas?", "Casi sin melena", "Con una melena enorme", "Con rayas como un tigre"],
  ["paleolitico", 14, "¿Con qué iluminaban las cuevas?", "Con lámparas de piedra que quemaban grasa", "Con velas de cera", "Con luciérnagas"],
  ["paleolitico", 15, "¿Qué hacen los osos pardos en invierno?", "Dormir dentro de las cuevas", "Viajar al sur", "Pescar en el hielo"],
  ["paleolitico", 16, "¿Cómo pintaban las manos en las cuevas?", "Soplando ocre alrededor de la mano", "Con un pincel de pelo", "Con témperas de colores"],
  ["neolitico", 1, "¿Dónde está el poblado de La Draga?", "Junto al lago de Banyoles (Girona)", "En una isla del Mediterráneo", "En lo alto de una montaña"],
  ["neolitico", 4, "¿Con qué trabajaban la tierra?", "Con palos cavadores de madera", "Con un tractor", "Con un arado de hierro"],
  ["neolitico", 7, "¿Cuál fue el primer animal domesticado?", "El perro", "La gallina", "El caballo"],
  ["neolitico", 10, "¿Qué hacían con el grano molido?", "Gachas y tortas", "Pizza", "Palomitas"],
  ["neolitico", 13, "¿Cómo se decoraba la cerámica cardial?", "Con el borde de una concha", "Con pintura dorada", "Con pegatinas"],
  ["neolitico", 14, "¿De qué madera eran los arcos de La Draga?", "De tejo", "De bambú", "De pino"],
  ["neolitico", 15, "¿Por qué secaban las vasijas antes de cocerlas?", "Para que no explotaran en el fuego", "Para pintarlas mejor", "Porque olían mal"],
  ["neolitico", 18, "¿Qué significa «Neolítico»?", "Piedra nueva", "Piedra antigua", "Edad del barro"],
  ["neolitico", 23, "¿Cómo eran las hoces?", "De madera con dientes de sílex", "De hierro", "De bronce"],
  ["neolitico", 30, "¿Por qué se conservó la madera de La Draga?", "Porque quedó bajo el agua y el barro", "Porque la barnizaron", "Porque era de piedra"]
];

/** Cambia una celda buscando la fila por una columna clave. Solo si está vacía (no pisa lo editado a mano). */
function _rellenar(tab, keyCol, keyVal, col, val) {
  const sh = _ss().getSheetByName(tab), vals = sh.getDataRange().getValues(), head = vals[0].map(String);
  const k = head.indexOf(keyCol), c = head.indexOf(col);
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][k]) === keyVal && (vals[r][c] === '' || vals[r][c] === null)) sh.getRange(r + 1, c + 1).setValue(val);
  }
}
/** Añade lo nuevo de la v0.6 sin tocar lo que ya hay. Idempotente (lo llama setup). */
function _completarV06() {
  const avs = _readSheet('Aventuras').map(a => String(a.id));
  _rellenar('Aventuras', 'id', 'paleolitico', 'hace', 15000);
  if (avs.indexOf('neolitico') < 0) {
    _append('Aventuras', NEO_AVENTURA);
    NEO_ESCENAS.forEach(r => _append('Escenas', { aventura_id: 'neolitico', num: r[0], tipo: r[1], titulo: r[2], texto: r[3], dato: r[4],
      opcion_a: r[5], destino_a: r[6] || '', opcion_b: r[7], destino_b: r[8] || '', emoji: r[9], sonido: r[10],
      imagen: 'neolitico_' + ('0' + r[0]).slice(-2) + '.webp' }));
  }
  if (avs.indexOf('metales') < 0) _append('Aventuras', METALES_PRONTO);
  if (!_readSheet('Quiz').length) QUIZ_SEMILLA.forEach(q => _append('Quiz', {
    aventura_id: q[0], escena: q[1], pregunta: q[2], correcta: q[3], incorrecta_1: q[4], incorrecta_2: q[5] }));
}

/** v0.9: categorías. Las de historia se marcan solas; las temáticas de ciencia entran como «pronto». */
function _completarV09() {
  ['paleolitico', 'neolitico', 'metales'].forEach(id => _rellenar('Aventuras', 'id', id, 'categoria', 'historia'));
  const avs = _readSheet('Aventuras').map(a => String(a.id));
  [
    { id: 'cuerpo', titulo: 'El cuerpo humano', epoca: 'Un viaje por dentro de ti', emoji: '🫀', color: '#F2A7A0' },
    { id: 'oceano', titulo: 'El océano', epoca: 'Hasta el fondo del mar', emoji: '🐙', color: '#9CCBE8' },
    { id: 'espacio', titulo: 'El espacio', epoca: 'Misión a la Luna', emoji: '🚀', color: '#C8B6E2' }
  ].forEach((a, i) => { if (avs.indexOf(a.id) < 0) _append('Aventuras', Object.assign({ categoria: 'ciencia', activa: 'pronto', orden: 10 + i }, a)); });
}

/* ---------- v0.10: Edad de los Metales (con mochila) ---------- */
const METALES_AVENTURA = {"id": "metales", "titulo": "Edad de los Metales", "epoca": "Hace 4.000 años · El Argar (sureste)", "intro": "Vives en un poblado de El Argar, en lo alto de un cerro. El abuelo, el fundidor, se ha torcido un tobillo… y te toca a ti fabricar tu primera hacha de bronce. ¡Llena bien tu mochila!", "color": "#E8C48A", "emoji": "⚒️", "kit": "metales", "hace": 4000, "categoria": "historia", "activa": true, "orden": 3, "imagen": "metales_portada.webp", "quiz_n": 10};
const METALES_OBJETOS = [{"aventura_id": "metales", "id": "plata", "nombre": "Brazalete de plata", "emoji": "💍", "descripcion": "El brazalete favorito de la abuela. Brilla muchísimo.", "dato": "La plata era muy valiosa. En los poblados de El Argar se han encontrado brazaletes, anillos y diademas de plata."}, {"aventura_id": "metales", "id": "manta", "nombre": "Manta de lana", "emoji": "🧣", "descripcion": "La manta que tejió mamá. Calentita y suave.", "dato": "La lana de oveja se hilaba y se tejía en telares para hacer ropa y mantas."}, {"aventura_id": "metales", "id": "malaquita", "nombre": "Piedra verde", "emoji": "🟢", "descripcion": "Una piedra verde, sin brillo, de la mina.", "dato": "Es malaquita, un mineral de cobre. Al fundirla en el horno sale cobre."}, {"aventura_id": "metales", "id": "cuarzo", "nombre": "Piedra blanca brillante", "emoji": "💎", "descripcion": "Una piedra blanca que brilla muchísimo.", "dato": "Es cuarzo: brilla mucho, pero no tiene metal dentro."}, {"aventura_id": "metales", "id": "carbon", "nombre": "Saco de carbón", "emoji": "⚫", "descripcion": "Carbón vegetal de los carboneros del bosque.", "dato": "El carbón vegetal arde mucho más fuerte que la leña: por eso se usaba en los hornos para fundir."}, {"aventura_id": "metales", "id": "cobre", "nombre": "Lingote de cobre", "emoji": "🟠", "descripcion": "Un lingote rojizo que acabas de fundir.", "dato": "El cobre fue uno de los primeros metales que se trabajaron. Es bastante blando."}, {"aventura_id": "metales", "id": "estano", "nombre": "Lingote de estaño", "emoji": "⚪", "descripcion": "Un lingote plateado que trae el mercader.", "dato": "El estaño era raro por aquí y llegaba de tierras lejanas. Mezclado con cobre forma el bronce."}, {"aventura_id": "metales", "id": "molde", "nombre": "Molde de piedra", "emoji": "🪨", "descripcion": "El molde del abuelo, con forma de hacha.", "dato": "Los fundidores tallaban moldes de piedra y vertían dentro el metal fundido."}, {"aventura_id": "metales", "id": "hacha", "nombre": "Hacha de bronce", "emoji": "🪓", "descripcion": "¡Tu primera hacha de bronce!", "dato": "El bronce es más duro que el cobre: con él se hacían hachas, puñales, punzones y hoces."}];
const METALES_ESCENAS = [{"aventura_id": "metales", "num": 1, "tipo": "decision", "titulo": "El poblado del cerro", "texto": "Vives en un poblado de casas de piedra y barro, en lo alto de un cerro. El abuelo, el fundidor, se ha torcido un tobillo… ¡y pronto llega el mercader! Te pide que hagas tú el hacha de bronce. Antes de salir, puedes meter UNA cosa en la mochila.", "dato": "En la Edad de los Metales los poblados crecieron, y muchos se construían en lugares altos y protegidos.", "sonido": "", "obtiene": "", "imagen": "metales_01.webp", "opcion_a": "El brazalete de plata de la abuela", "destino_a": 2, "obtiene_a": "plata", "opcion_b": "La manta de lana que tejió mamá", "destino_b": 2, "obtiene_b": "manta"}, {"aventura_id": "metales", "num": 2, "tipo": "decision", "titulo": "Las instrucciones", "texto": "El abuelo te explica: «Para fundir cobre necesito mineral de la mina y carbón del bosque. El estaño lo conseguiremos con el mercader». ¿Por dónde empiezas?", "dato": "El mineral de cobre más fácil de reconocer es la malaquita: una piedra de color verde.", "sonido": "", "obtiene": "", "imagen": "metales_02.webp", "opcion_a": "Subir a la mina de la montaña", "destino_a": 3, "opcion_b": "Ir al bosque de los carboneros", "destino_b": 4}, {"aventura_id": "metales", "num": 3, "tipo": "decision", "titulo": "La mina", "texto": "En la mina, los mineros golpean la roca. En el suelo ves dos piedras: una blanca que brilla muchísimo y otra verde, sin brillo. Solo cabe una en la mochila.", "dato": "Los mineros sacaban el mineral golpeando la roca con mazas de piedra y picos hechos con asta de ciervo.", "sonido": "", "obtiene": "", "imagen": "metales_03.webp", "opcion_a": "Coger la piedra blanca y brillante", "destino_a": 5, "obtiene_a": "cuarzo", "opcion_b": "Coger la piedra verde", "destino_b": 5, "obtiene_b": "malaquita"}, {"aventura_id": "metales", "num": 4, "tipo": "decision", "titulo": "Los carboneros", "texto": "En el bosque, los carboneros amontonan leña, la encienden y la tapan con tierra. Te darán un saco de carbón si les ayudas con la pila.", "dato": "Para hacer carbón vegetal, la leña se quema muy despacio, tapada con tierra y casi sin aire.", "sonido": "", "obtiene": "", "imagen": "metales_04.webp", "opcion_a": "Echar más tierra encima de la pila", "destino_a": 10, "obtiene_a": "carbon", "opcion_b": "Abrir un agujero para que respire el fuego", "destino_b": 9}, {"aventura_id": "metales", "num": 5, "tipo": "decision", "titulo": "Se hace de noche", "texto": "Con la piedra en la mochila, el sol se esconde. En la montaña hace mucho frío y el camino de bajada está muy oscuro.", "dato": "Hace 4.000 años no había farolas: de noche solo alumbraban la luna y el fuego.", "sonido": "", "obtiene": "", "imagen": "metales_05.webp", "opcion_a": "Quedarte a dormir con los mineros", "destino_a": 7, "requiere_a": "manta", "sino_a": 6, "opcion_b": "Bajar por el sendero de las cabras", "destino_b": 8}, {"aventura_id": "metales", "num": 6, "tipo": "fin", "titulo": "¡Tiritona!", "texto": "Sin nada con que abrigarte, pasas la noche temblando como un flan. Por la mañana estornudas tan fuerte que espantas a todas las cabras de la montaña.", "dato": "La ropa y las mantas de lana eran imprescindibles para pasar las noches frías.", "sonido": "¡ACHÚS!", "obtiene": "", "imagen": "metales_06.webp"}, {"aventura_id": "metales", "num": 7, "tipo": "decision", "titulo": "Noche con los mineros", "texto": "Envuelta en tu manta, cenas junto al fuego de los mineros. Al amanecer, uno te regala un saco de carbón: «Para el horno de tu abuelo».", "dato": "Mineros, fundidores, alfareros, tejedoras… en esta época cada vez más gente se dedicaba a un oficio.", "sonido": "", "obtiene": "carbon", "imagen": "metales_07.webp", "opcion_a": "Bajar al taller del abuelo", "destino_a": 11, "opcion_b": "Pasar antes por el bosque de los carboneros", "destino_b": 4}, {"aventura_id": "metales", "num": 8, "tipo": "decision", "titulo": "De vuelta al poblado", "texto": "Llegas al poblado de noche, muerta de sueño. Por la mañana, ¿adónde vas?", "dato": "Las casas solían tener un banco para dormir, un hogar para el fuego y vasijas para guardar el grano.", "sonido": "", "obtiene": "", "imagen": "metales_08.webp", "opcion_a": "Al bosque de los carboneros", "destino_a": 4, "opcion_b": "Al taller del abuelo", "destino_b": 11}, {"aventura_id": "metales", "num": 9, "tipo": "fin", "titulo": "¡Fuuush!", "texto": "Entra aire, la pila se convierte en una llamarada… y en vez de carbón queda un montón de ceniza. Los carboneros te persiguen con las cejas chamuscadas.", "dato": "Con mucho aire la madera se quema entera y solo deja ceniza. Con poco aire se convierte en carbón.", "sonido": "¡FUUUSH!", "obtiene": "", "imagen": "metales_09.webp"}, {"aventura_id": "metales", "num": 10, "tipo": "decision", "titulo": "Un saco de carbón", "texto": "¡Lo conseguiste! Con el saco de carbón al hombro, piensas en lo que te falta.", "dato": "Hacían falta muchos sacos de carbón para fundir un poco de metal.", "sonido": "", "obtiene": "", "imagen": "metales_10.webp", "opcion_a": "Subir a la mina a por mineral", "destino_a": 3, "opcion_b": "Ir directa al taller del abuelo", "destino_b": 11}, {"aventura_id": "metales", "num": 11, "tipo": "decision", "titulo": "El taller del abuelo", "texto": "En el taller hay un horno de barro, un crisol y un fuelle de piel. El abuelo, con el pie vendado, te pregunta: «¿Lo tienes todo?».", "dato": "Para fundir el cobre hace falta muchísimo más calor que en una hoguera: avivaban el carbón soplando aire con fuelles.", "sonido": "", "obtiene": "", "imagen": "metales_11.webp", "opcion_a": "¡Sí! Fundir el mineral", "destino_a": 15, "requiere_a": "malaquita,carbon", "sino_a": 12, "gasta_a": "malaquita,carbon", "opcion_b": "Todavía no: salir a buscar lo que falta", "destino_b": 14, "opcion_c": "Fundir la piedra blanca y brillante", "destino_c": 13, "requiere_c": "cuarzo", "gasta_c": "cuarzo"}, {"aventura_id": "metales", "num": 12, "tipo": "decision", "titulo": "Así no sale nada", "texto": "Soplas y soplas con el fuelle… pero no sale ni una gota de metal. El abuelo mira dentro de tu mochila: «Te falta algo».", "dato": "Para obtener cobre hace falta mineral de cobre y carbón, las dos cosas juntas.", "sonido": "", "obtiene": "", "imagen": "metales_12.webp", "opcion_a": "Ir a la mina", "destino_a": 3, "opcion_b": "Ir al bosque de los carboneros", "destino_b": 4}, {"aventura_id": "metales", "num": 13, "tipo": "fin", "titulo": "Piedra asada", "texto": "La piedra blanca se calienta, se calienta… y sigue siendo una piedra blanca, ¡pero muy caliente! El abuelo se ríe: «Eso no tiene ni pizca de metal».", "dato": "Era cuarzo: brilla mucho, pero no lleva cobre dentro. ¡No todo lo que brilla es metal!", "sonido": "¡OPS!", "obtiene": "", "imagen": "metales_13.webp"}, {"aventura_id": "metales", "num": 14, "tipo": "decision", "titulo": "¿Qué te falta?", "texto": "Abres la mochila y revisas lo que llevas. ¿Adónde vas?", "dato": "Los fundidores sabían reconocer los minerales por su color y su peso.", "sonido": "", "obtiene": "", "imagen": "metales_14.webp", "opcion_a": "A la mina", "destino_a": 3, "opcion_b": "Al bosque de los carboneros", "destino_b": 4}, {"aventura_id": "metales", "num": 15, "tipo": "decision", "titulo": "¡Cobre!", "texto": "¡El mineral se funde y sale cobre líquido, rojo como el fuego! El abuelo te da un lingote de cobre y su molde de piedra con forma de hacha.", "dato": "El cobre fue uno de los primeros metales que se trabajaron. Es bonito, pero bastante blando.", "sonido": "", "obtiene": "cobre,molde", "imagen": "metales_15.webp", "opcion_a": "Hacer ya un hacha de cobre", "destino_a": 16, "opcion_b": "Esperar al mercader para conseguir estaño", "destino_b": 17, "opcion_c": "¡Mezclarlo con el estaño que llevas!", "destino_c": 24, "requiere_c": "estano"}, {"aventura_id": "metales", "num": 16, "tipo": "fin", "titulo": "Hacha blandita", "texto": "Haces un hacha de cobre preciosa. Al primer golpe contra un tronco de encina… ¡el filo se dobla! Tu compi se ríe tanto que se cae de culo.", "dato": "El cobre es blando. Mezclado con un poco de estaño se convierte en bronce, mucho más duro.", "sonido": "¡BOING!", "obtiene": "", "imagen": "metales_16.webp"}, {"aventura_id": "metales", "num": 17, "tipo": "decision", "titulo": "Llega el mercader", "texto": "El mercader llega a la plaza con un fardo a la espalda. ¡Trae estaño! Pero no lo regala: quiere algo a cambio.", "dato": "El estaño era muy raro por aquí: llegaba de tierras lejanas gracias al intercambio entre pueblos.", "sonido": "", "obtiene": "", "imagen": "metales_17.webp", "opcion_a": "Ofrecerle tu lingote de cobre", "destino_a": 18, "obtiene_a": "estano", "gasta_a": "cobre", "opcion_b": "Ofrecerle tu manta de lana", "destino_b": 24, "requiere_b": "manta", "sino_b": 21, "obtiene_b": "estano", "gasta_b": "manta", "opcion_c": "Ofrecerle el brazalete de plata", "destino_c": 24, "requiere_c": "plata", "obtiene_c": "estano", "gasta_c": "plata"}, {"aventura_id": "metales", "num": 18, "tipo": "decision", "titulo": "¿Y el cobre?", "texto": "¡Tienes estaño! Pero el abuelo frunce el ceño: «¿Y con qué lo vas a mezclar, si le has dado tu cobre?».", "dato": "El bronce es una aleación: una mezcla de cobre y estaño. Sin cobre no hay bronce.", "sonido": "", "obtiene": "", "imagen": "metales_18.webp", "opcion_a": "Hacer el hacha solo con estaño", "destino_a": 20, "opcion_b": "Volver a la mina a por más mineral", "destino_b": 19}, {"aventura_id": "metales", "num": 19, "tipo": "decision", "titulo": "Otra vez a la mina", "texto": "Subes otra vez a la mina. Los mineros están sujetando un túnel con troncos, pero te dan un trozo de piedra verde y te mandan a casa.", "dato": "Las galerías de las minas se sujetaban con troncos para que no se hundieran.", "sonido": "", "obtiene": "malaquita", "imagen": "metales_19.webp", "opcion_a": "Pedir carbón a los carboneros", "destino_a": 4, "opcion_b": "Ir directa al taller", "destino_b": 11}, {"aventura_id": "metales", "num": 20, "tipo": "fin", "titulo": "Hacha de plastilina", "texto": "El estaño se funde enseguida y haces un hacha brillante… tan blanda que se dobla con los dedos. ¡Parece de plastilina!", "dato": "El estaño solo es muy blando. Mezclado con cobre, en cambio, forma el bronce, que es duro.", "sonido": "¡ÑIIIC!", "obtiene": "", "imagen": "metales_20.webp"}, {"aventura_id": "metales", "num": 21, "tipo": "decision", "titulo": "¡No la llevas!", "texto": "Buscas la manta en la mochila… ¡pero no está! El mercader se cruza de brazos y espera.", "dato": "En esta época no había monedas: se cambiaba una cosa por otra.", "sonido": "", "obtiene": "", "imagen": "metales_21.webp", "opcion_a": "Ofrecerle el lingote de cobre", "destino_a": 18, "obtiene_a": "estano", "gasta_a": "cobre", "opcion_b": "Ofrecerte a cargar su fardo hasta el río", "destino_b": 22, "opcion_c": "Ofrecerle el brazalete de plata", "destino_c": 24, "requiere_c": "plata", "obtiene_c": "estano", "gasta_c": "plata"}, {"aventura_id": "metales", "num": 22, "tipo": "decision", "titulo": "El fardo del mercader", "texto": "El fardo pesa como un ternero. El mercader te dará un poco de estaño si lo llevas hasta el río.", "dato": "Las mercancías se transportaban a pie, a lomos de animales o en barcas por los ríos.", "sonido": "", "obtiene": "", "imagen": "metales_22.webp", "opcion_a": "Cargarlo tú sola, ¡puedes con todo!", "destino_a": 23, "opcion_b": "Pedir ayuda a tu compi y llevarlo entre dos", "destino_b": 24, "obtiene_b": "estano"}, {"aventura_id": "metales", "num": 23, "tipo": "fin", "titulo": "Aplastada", "texto": "Das tres pasos… y el fardo te aplasta como una tortilla. El mercader tiene que sacarte de debajo tirando de los pies.", "dato": "Entre varias personas se podían mover cargas muy pesadas.", "sonido": "¡CHOF!", "obtiene": "", "imagen": "metales_23.webp"}, {"aventura_id": "metales", "num": 24, "tipo": "decision", "titulo": "La mezcla", "texto": "Con el cobre y el estaño en la mochila, vuelves al taller. El abuelo pregunta: «¿Cuánto estaño ponemos?».", "dato": "En el bronce se ponía más o menos una parte de estaño por cada nueve de cobre.", "sonido": "", "obtiene": "", "imagen": "metales_24.webp", "opcion_a": "Mitad y mitad, ¡que quede fortísimo!", "destino_a": 25, "opcion_b": "Un poquito de estaño y mucho cobre", "destino_b": 26, "requiere_b": "cobre,estano", "sino_b": 18, "gasta_b": "cobre,estano"}, {"aventura_id": "metales", "num": 25, "tipo": "fin", "titulo": "Bronce de galleta", "texto": "El bronce sale precioso, pero al primer golpe… ¡se parte como una galleta! Con tanto estaño, el bronce se vuelve quebradizo.", "dato": "Con demasiado estaño el bronce se rompe; con la mezcla justa es duro y resistente.", "sonido": "¡CRAC!", "obtiene": "", "imagen": "metales_25.webp"}, {"aventura_id": "metales", "num": 26, "tipo": "decision", "titulo": "¡A verter!", "texto": "El bronce líquido brilla dentro del crisol. Ahora hay que echarlo en el molde.", "dato": "El metal fundido se enfría y se endurece muy rápido: hay que verterlo enseguida.", "sonido": "", "obtiene": "", "imagen": "metales_26.webp", "opcion_a": "Esperar un poco a que no queme tanto", "destino_a": 27, "opcion_b": "Verterlo enseguida en el molde", "destino_b": 28, "requiere_b": "molde", "sino_b": 27}, {"aventura_id": "metales", "num": 27, "tipo": "fin", "titulo": "Bronce en el crisol", "texto": "Esperas… y el bronce se endurece dentro del crisol. Ahora tienes un pegote de bronce con forma de cazo. ¡El abuelo lo usa para cascar nueces!", "dato": "Por eso los fundidores trabajaban muy deprisa al verter el metal.", "sonido": "¡TOC!", "obtiene": "", "imagen": "metales_27.webp"}, {"aventura_id": "metales", "num": 28, "tipo": "decision", "titulo": "Tu hacha de bronce", "texto": "¡Del molde sale un hacha de bronce! Todavía tiene el filo grueso. ¿Qué haces?", "dato": "Después de sacarlas del molde, golpeaban el filo con un martillo para endurecerlo y lo afilaban frotándolo con piedras.", "sonido": "", "obtiene": "hacha", "imagen": "metales_28.webp", "opcion_a": "Golpear el filo con el martillo y afilarlo con una piedra", "destino_a": 31, "opcion_b": "Probarla enseguida contra una roca", "destino_b": 32}, {"aventura_id": "metales", "num": 29, "tipo": "final", "titulo": "Maestra del bronce", "texto": "Tu hacha corta la encina como si fuera mantequilla. Esa noche todo el poblado lo celebra y el abuelo presume: «¡Tenemos una nueva fundidora!».", "dato": "En la Edad de los Metales se extendieron grandes inventos como la rueda, el arado y la vela.", "sonido": "¡HURRA!", "obtiene": "", "imagen": "metales_29.webp"}, {"aventura_id": "metales", "num": 30, "tipo": "final", "titulo": "La gran comerciante", "texto": "Cambias el hacha al mercader por un buen montón de estaño. Durante el invierno, el abuelo y tú hacéis hachas, punzones y hoces para todo el poblado.", "dato": "El intercambio de productos entre pueblos hizo crecer el comercio.", "sonido": "¡HURRA!", "obtiene": "", "imagen": "metales_30.webp"}, {"aventura_id": "metales", "num": 31, "tipo": "decision", "titulo": "¿Y ahora?", "texto": "¡El hacha queda dura y afilada! El mercader la mira con ojos golosos: «Te la cambio por estaño para hacer muchas más».", "dato": "Más adelante llegaría el hierro: se encontraba en muchos más sitios que el cobre y el estaño.", "sonido": "", "obtiene": "", "imagen": "metales_31.webp", "opcion_a": "Quedártela para el poblado", "destino_a": 29, "opcion_b": "Cambiarla por el estaño del mercader", "destino_b": 30, "gasta_b": "hacha"}, {"aventura_id": "metales", "num": 32, "tipo": "fin", "titulo": "Mellada", "texto": "¡CLANC! El filo se mella contra la roca y salta una chispa que casi le quema el bigote al abuelo. ¡Un hacha es para la madera, no para las piedras!", "dato": "Las herramientas de metal se cuidaban mucho: eran valiosas y costaba mucho hacerlas.", "sonido": "¡CLANC!", "obtiene": "", "imagen": "metales_32.webp"}];
const METALES_QUIZ = [["metales", 0, 1, "¿Qué aprendieron a trabajar las personas en la Edad de los Metales?", "Los metales", "El plástico", "El papel"], ["metales", 0, 1, "¿Qué dos metales se mezclan para hacer bronce?", "Cobre y estaño", "Hierro y plata", "Oro y cobre"], ["metales", 15, 1, "¿Cuál fue uno de los primeros metales que se trabajaron?", "El cobre", "El hierro", "El aluminio"], ["metales", 2, 1, "¿De qué color es la malaquita, un mineral de cobre?", "Verde", "Blanco", "Negro"], ["metales", 0, 1, "¿Qué invento facilitó el transporte de mercancías por tierra?", "La rueda", "La aguja de hueso", "La cerámica"], ["metales", 0, 2, "¿Qué es la metalurgia?", "El trabajo de los metales para fabricar objetos", "El estudio de las estrellas", "La cría de animales"], ["metales", 0, 2, "¿Para qué servía el arado?", "Para trabajar la tierra con más facilidad", "Para navegar por el mar", "Para fundir metales"], ["metales", 0, 2, "¿Para qué servía la vela en las barcas?", "Para aprovechar el viento", "Para iluminar de noche", "Para pescar"], ["metales", 11, 2, "¿Por qué usaban carbón y fuelles para fundir el cobre?", "Porque hace falta muchísimo calor", "Para que oliera bien", "Para pintar el metal de negro"], ["metales", 0, 2, "¿En qué orden se empezaron a trabajar estos metales?", "Cobre, bronce, hierro", "Hierro, bronce, cobre", "Bronce, hierro, cobre"], ["metales", 21, 2, "Sin monedas, ¿cómo comerciaban?", "Cambiando unas cosas por otras", "Con billetes de papel", "Con tarjetas"], ["metales", 0, 3, "¿Qué es una aleación?", "Una mezcla de dos o más elementos, al menos uno de ellos un metal", "Un metal sin mezclar", "Un tipo de horno"], ["metales", 24, 3, "¿Cuánto estaño se ponía, más o menos, en el bronce?", "Una parte por cada nueve de cobre", "Mitad y mitad", "Más estaño que cobre"], ["metales", 17, 3, "¿Por qué el estaño era tan valioso en el sureste de la Península?", "Porque llegaba de tierras lejanas", "Porque se podía comer", "Porque brillaba en la oscuridad"], ["metales", 0, 3, "¿Qué cambió en los poblados de la Edad de los Metales?", "Crecieron y aparecieron distintos oficios", "Volvieron a ser nómadas", "Dejaron de cultivar"], ["metales", 15, 3, "¿Por qué el bronce era mejor que el cobre para hacer herramientas?", "Porque es más duro y resistente", "Porque flota en el agua", "Porque nunca se calienta"]];

/** Activa la Edad de los Metales: rellena su fila (estaba como «pronto») y carga escenas, objetos y quiz
    si todavía no existen. No toca nada que ya esté. */
function _completarV10() {
  const sh = _ss().getSheetByName('Aventuras'), vals = sh.getDataRange().getValues(), head = vals[0].map(String);
  const iId = head.indexOf('id');
  let fila = -1;
  for (let r = 1; r < vals.length; r++) if (String(vals[r][iId]) === 'metales') fila = r + 1;
  const tieneEscenas = _readSheet('Escenas').some(e => String(e.aventura_id) === 'metales');
  if (fila < 0) _append('Aventuras', METALES_AVENTURA);
  else if (!tieneEscenas) Object.keys(METALES_AVENTURA).forEach(k => {
    const c = head.indexOf(k); if (c >= 0) sh.getRange(fila, c + 1).setValue(METALES_AVENTURA[k]);
  });
  if (!tieneEscenas) METALES_ESCENAS.forEach(e => _append('Escenas', e));
  if (!_readSheet('Objetos').some(o => String(o.aventura_id) === 'metales')) METALES_OBJETOS.forEach(o => _append('Objetos', o));
  if (!_readSheet('Quiz').some(q => String(q.aventura_id) === 'metales')) METALES_QUIZ.forEach(q => _append('Quiz', {
    aventura_id: q[0], escena: q[1], nivel: q[2], pregunta: q[3], correcta: q[4], incorrecta_1: q[5], incorrecta_2: q[6] }));
}

/** v0.11: Gina no tiene hermanos → «tu hermano» pasa a «tu compi» en las escenas ya cargadas. */
function _completarV11() {
  const sh = _ss().getSheetByName('Escenas'), rg = sh.getDataRange(), vals = rg.getValues();
  let n = 0;
  for (let r = 1; r < vals.length; r++) for (let c = 0; c < vals[r].length; c++) {
    const v = vals[r][c];
    if (typeof v === 'string' && /\b[Tt]u hermano\b/.test(v)) { sh.getRange(r + 1, c + 1).setValue(v.replace(/\b([Tt])u hermano\b/g, '$1u compi')); n++; }
  }
  if (n) Logger.log('«hermano» → «compi» en ' + n + ' celdas');
}

/* ---------- v0.12: De renacuajo a rana (Naturaleza) ---------- */
const RANA_AVENTURA = {"id": "rana", "titulo": "De renacuajo a rana", "epoca": "Una charca · ciclo de vida", "intro": "Papá te enseña los huevos de rana de la charca del jardín. Te quedas dormida mirándolos… ¡y sueñas que eres un renacuajo! Crece, esquiva peligros y conviértete en rana. Lo que te va saliendo en el cuerpo te abrirá caminos.", "color": "#8CC7A1", "emoji": "🐸", "kit": "rana", "categoria": "naturaleza", "activa": true, "orden": 1, "imagen": "rana_portada.webp", "quiz_n": 10, "mochila": "Tu cuerpo", "mochila_emoji": "🐸"};
const RANA_OBJETOS = [{"aventura_id": "rana", "id": "branquias", "nombre": "Branquias", "emoji": "🫧", "descripcion": "Para respirar bajo el agua, como los peces.", "dato": "Los renacuajos respiran por branquias. Al convertirse en ranas, respiran con pulmones y también a través de la piel."}, {"aventura_id": "rana", "id": "cola", "nombre": "Cola", "emoji": "〰️", "descripcion": "Tu cola para nadar. ¡Zas, zas!", "dato": "La cola no se cae: el cuerpo la va absorbiendo poco a poco y la usa como alimento."}, {"aventura_id": "rana", "id": "patas_traseras", "nombre": "Patas traseras", "emoji": "🦵", "descripcion": "Dos patas fuertes para impulsarte.", "dato": "En la metamorfosis salen primero las patas traseras."}, {"aventura_id": "rana", "id": "patas_delanteras", "nombre": "Patas delanteras", "emoji": "✋", "descripcion": "Dos patitas cortas para apoyarte.", "dato": "Las patas delanteras salen después de las traseras."}, {"aventura_id": "rana", "id": "pulmones", "nombre": "Pulmones", "emoji": "🫁", "descripcion": "Para respirar aire fuera del agua.", "dato": "Con pulmones, la rana puede vivir fuera del agua, aunque necesita tener la piel húmeda."}, {"aventura_id": "rana", "id": "lengua", "nombre": "Lengua pegajosa", "emoji": "👅", "descripcion": "Larga, rápida y pegajosa. ¡Ideal para cazar!", "dato": "La rana dispara su lengua pegajosa en un instante para atrapar insectos."}];
const RANA_ESCENAS = [{"aventura_id": "rana", "num": 1, "tipo": "decision", "titulo": "Un sueño en la charca", "texto": "Papá te ha enseñado los huevos de rana de la charca del jardín. Te tumbas a mirarlos… y te quedas dormida. ¡Sueñas que acabas de salir de un huevo! Eres un renacuajo: tienes cola y respiras bajo el agua. A tu lado nadan dos renacuajos que te suenan muchísimo: ¡Lola y Diana!", "dato": "Al nacer, los renacuajos todavía nadan muy mal: los primeros días se quedan pegados a las plantas.", "sonido": "", "obtiene": "branquias,cola", "imagen": "rana_01.webp", "opcion_a": "Agarrarte a una planta, como los demás", "destino_a": 3, "opcion_b": "¡Nadar a explorar toda la charca!", "destino_b": 2}, {"aventura_id": "rana", "num": 2, "tipo": "fin", "titulo": "Merienda de pez", "texto": "Nadas haciendo eses, sin rumbo… y un pez te confunde con un aperitivo. Por suerte le das tanto asco que te escupe. ¡PTUÉ!", "dato": "Muchos peces comen renacuajos. Por eso las ranas ponen cientos de huevos: así algunos llegan a adultos.", "sonido": "¡PTUÉ!", "obtiene": "", "imagen": "rana_02.webp"}, {"aventura_id": "rana", "num": 3, "tipo": "decision", "titulo": "¡Qué hambre!", "texto": "Ya nadas mucho mejor y te ruge la barriguita. Diana ya está rascando las piedras, cubiertas de algas verdes. Sobre el agua zumba un mosquito.", "dato": "Los renacuajos comen sobre todo algas y restos de plantas: los rascan con su boquita.", "sonido": "", "obtiene": "", "imagen": "rana_03.webp", "opcion_a": "Saltar a por el mosquito", "destino_a": 4, "opcion_b": "Rascar las algas de las piedras", "destino_b": 5}, {"aventura_id": "rana", "num": 4, "tipo": "fin", "titulo": "Salto imposible", "texto": "Saltas a por el mosquito… pero no tienes patas ni lengua para cazar. Te quedas flotando panza arriba mientras el mosquito se ríe.", "dato": "Los renacuajos no pueden cazar insectos que vuelan: eso lo harán de ranas, con su lengua.", "sonido": "¡BZZZ!", "obtiene": "", "imagen": "rana_04.webp"}, {"aventura_id": "rana", "num": 5, "tipo": "decision", "titulo": "Una sombra en el fondo", "texto": "Mientras comes, algo se acerca muy despacio por el fondo: un bicho alargado con unas mandíbulas raras.", "dato": "Es una ninfa de libélula: vive bajo el agua y caza renacuajos. Los renacuajos se salvan escondiéndose entre las plantas.", "sonido": "", "obtiene": "", "imagen": "rana_05.webp", "opcion_a": "Quedarte quieta: parece muy tranquila", "destino_a": 6, "opcion_b": "Esconderte entre las plantas más espesas", "destino_b": 7}, {"aventura_id": "rana", "num": 6, "tipo": "fin", "titulo": "¡ÑAC!", "texto": "Su mandíbula sale disparada como un brazo extensible. ¡ÑAC! Te escapas por los pelos… pero sin un trocito de cola y con un susto que no se te quita.", "dato": "La ninfa de libélula tiene una mandíbula plegada que se dispara hacia delante para atrapar presas.", "sonido": "¡ÑAC!", "obtiene": "", "imagen": "rana_06.webp"}, {"aventura_id": "rana", "num": 7, "tipo": "decision", "titulo": "¡Patitas!", "texto": "Escondida entre las plantas pasan las semanas… ¡y te salen dos patitas traseras! A Diana también. Lola es un poco más pequeñita: a ella todavía no le han salido. El sol calienta la orilla.", "dato": "En la metamorfosis salen primero las patas traseras y después las delanteras.", "sonido": "", "obtiene": "patas_traseras", "imagen": "rana_07.webp", "opcion_a": "Nadar al centro, donde hay más comida", "destino_a": 8, "opcion_b": "Salir a la orilla a tomar el sol", "destino_b": 13, "requiere_b": "patas_delanteras,pulmones", "sino_b": 15}, {"aventura_id": "rana", "num": 8, "tipo": "decision", "titulo": "Una estatua con pico", "texto": "En el centro de la charca hay mucha comida… y en la orilla, una garza quieta como una estatua, con el pico apuntando al agua.", "dato": "Las garzas cazan quietas, esperando a que algo se mueva cerca de la superficie.", "sonido": "", "obtiene": "", "imagen": "rana_08.webp", "opcion_a": "Subir a la superficie a mirarla mejor", "destino_a": 11, "opcion_b": "Bucear hasta el fondo, lejos de su pico", "destino_b": 9}, {"aventura_id": "rana", "num": 9, "tipo": "decision", "titulo": "El fondo de la charca", "texto": "En el fondo encuentras a un pez gordito y a un escarabajo buceador que brilla muchísimo. El escarabajo te mira con mucha atención.", "dato": "Los escarabajos buceadores son grandes cazadores: comen renacuajos y hasta pececillos.", "sonido": "", "obtiene": "", "imagen": "rana_09.webp", "opcion_a": "Seguir al escarabajo brillante", "destino_a": 12, "opcion_b": "Quedarte con el pez, junto al barro", "destino_b": 10, "opcion_c": "Alejarte nadando con tus patas nuevas", "destino_c": 10, "requiere_c": "patas_traseras"}, {"aventura_id": "rana", "num": 10, "tipo": "decision", "titulo": "Cuatro patas", "texto": "Pasa el tiempo junto al pez… ¡y te salen las patas delanteras! Tus branquias dejan paso a unos pulmones. El pez abre la boca: «Yo nunca tendré patas».", "dato": "Los peces respiran por branquias toda la vida. Las ranas son anfibios: viven en el agua de pequeñas y también fuera de mayores.", "sonido": "", "obtiene": "patas_delanteras,pulmones", "imagen": "rana_10.webp", "opcion_a": "Salir a la orilla a respirar aire", "destino_a": 13, "gasta_a": "branquias", "opcion_b": "Quedarte en el agua, que se está de maravilla", "destino_b": 14, "gasta_b": "branquias"}, {"aventura_id": "rana", "num": 11, "tipo": "fin", "titulo": "Picoteo", "texto": "Subes a mirar… ¡y la estatua se mueve! La garza te pesca con el pico. Te suelta del susto cuando Chewie le estornuda en la cara.", "dato": "Las garzas comen peces, ranas y renacuajos. Lo que más les llama la atención es el movimiento.", "sonido": "¡ZAS!", "obtiene": "", "imagen": "rana_11.webp"}, {"aventura_id": "rana", "num": 12, "tipo": "fin", "titulo": "Escarabajo glotón", "texto": "Sigues al escarabajo brillante… que se da la vuelta con cara de hambre. Sales nadando tan rápido que te mareas.", "dato": "Los escarabajos buceadores guardan una burbuja de aire bajo las alas para respirar bajo el agua.", "sonido": "¡GLUB!", "obtiene": "", "imagen": "rana_12.webp"}, {"aventura_id": "rana", "num": 13, "tipo": "decision", "titulo": "¡Ranita!", "texto": "Saltas a una hoja de nenúfar. Tu cola casi ha desaparecido: ¡eres una ranita! Una mosca se posa en una hoja cercana.", "dato": "De adultas, las ranas cambian de comida: dejan las algas y cazan insectos, gusanos y arañas.", "sonido": "", "obtiene": "lengua", "imagen": "rana_13.webp", "opcion_a": "Esperar y saltar encima de la mosca", "destino_a": 16, "opcion_b": "Lanzarle tu lengua pegajosa", "destino_b": 17, "requiere_b": "lengua", "sino_b": 16, "gasta_b": "cola"}, {"aventura_id": "rana", "num": 14, "tipo": "decision", "titulo": "La charca encoge", "texto": "Llega el verano. Hace mucho calor y el agua se va secando: la charca se hace pequeñita y el agua está tibia.", "dato": "En verano muchas charcas pequeñas se secan. Por eso los renacuajos tienen prisa por convertirse en ranas.", "sonido": "", "obtiene": "", "imagen": "rana_14.webp", "opcion_a": "Quedarte en el charquito que queda", "destino_a": 18, "opcion_b": "Saltar hasta la charca grande del bosque", "destino_b": 19, "requiere_b": "patas_delanteras,pulmones", "sino_b": 15}, {"aventura_id": "rana", "num": 15, "tipo": "fin", "titulo": "Pez fuera del agua", "texto": "Sales a la orilla… pero aún respiras por branquias. Boqueas como un pez fuera del agua hasta que vuelves de un salto, con la lengua fuera.", "dato": "Mientras respiran por branquias, los renacuajos no pueden vivir fuera del agua.", "sonido": "¡GLUP!", "obtiene": "", "imagen": "rana_15.webp"}, {"aventura_id": "rana", "num": 16, "tipo": "fin", "titulo": "Planchazo", "texto": "Saltas a por la mosca… fallas y caes al agua de panza. ¡PLAF! La mosca se limpia las patitas, muerta de risa.", "dato": "Las ranas no persiguen la comida: esperan quietas y disparan la lengua en un instante.", "sonido": "¡PLAF!", "obtiene": "", "imagen": "rana_16.webp"}, {"aventura_id": "rana", "num": 17, "tipo": "decision", "titulo": "¡Ñam!", "texto": "¡Atrapas la mosca con la lengua! Pasa el verano y llega el otoño. De pronto, una culebra de agua nada hacia ti.", "dato": "Las culebras de agua comen ranas y peces. Las ranas escapan saltando al agua y escondiéndose en el fondo.", "sonido": "", "obtiene": "", "imagen": "rana_17.webp", "opcion_a": "Hacerte la dormida en la hoja", "destino_a": 21, "opcion_b": "Saltar al agua y esconderte en el fondo", "destino_b": 20}, {"aventura_id": "rana", "num": 18, "tipo": "fin", "titulo": "Pegatina de barro", "texto": "El charquito se seca y te quedas pegada en el barro como una pegatina. Por suerte, una tormenta te rescata tres días después.", "dato": "Algunos anfibios aguantan en el barro húmedo, pero los renacuajos necesitan agua para vivir.", "sonido": "¡CHOF!", "obtiene": "", "imagen": "rana_18.webp"}, {"aventura_id": "rana", "num": 19, "tipo": "decision", "titulo": "La charca grande", "texto": "En la charca grande del bosque viven muchas ranas. Por la noche, algunas hinchan unas bolsas a los lados de la boca y croan: ¡CROAC, CROAC!", "dato": "Solo croan los machos: inflan unos sacos a los lados de la boca para que el sonido llegue lejos.", "sonido": "", "obtiene": "lengua", "imagen": "rana_19.webp", "opcion_a": "Probar a croar tú también", "destino_a": 22, "opcion_b": "Cazar mosquitos con tu lengua pegajosa", "destino_b": 25, "requiere_b": "lengua", "sino_b": 16, "gasta_b": "cola"}, {"aventura_id": "rana", "num": 20, "tipo": "decision", "titulo": "Llega el invierno", "texto": "Se acerca el invierno. El agua está helada, tu cuerpo va cada vez más lento y apenas hay insectos.", "dato": "Las ranas son de sangre fría: su cuerpo tiene la temperatura del lugar donde están. En invierno se esconden en el barro del fondo.", "sonido": "", "obtiene": "", "imagen": "rana_20.webp", "opcion_a": "Quedarte en la orilla, al solecito", "destino_a": 24, "opcion_b": "Enterrarte en el barro del fondo", "destino_b": 23}, {"aventura_id": "rana", "num": 21, "tipo": "fin", "titulo": "Siesta peligrosa", "texto": "Te haces la dormida… y la culebra se acerca a husmear. ¡Pegas un salto tan alto que aterrizas en la cabeza de un pato!", "dato": "Hacerse la muerta funciona con algunos cazadores, pero no con las culebras: ellas cazan por el olor.", "sonido": "¡CUAC!", "obtiene": "", "imagen": "rana_21.webp"}, {"aventura_id": "rana", "num": 22, "tipo": "fin", "titulo": "Gallito", "texto": "Hinchas los mofletes, respiras hondo… y te sale un «pío» de pollito. Todas las ranas de la charca se ríen. Lola y Diana, las que más.", "dato": "En la rana común solo croan los machos: las hembras son mucho más calladas.", "sonido": "¡PÍO!", "obtiene": "", "imagen": "rana_22.webp"}, {"aventura_id": "rana", "num": 23, "tipo": "final", "titulo": "¡Primavera!", "texto": "Duermes todo el invierno en el barro. Al despertar es primavera: la charca se llena de huevos nuevos… Y abres los ojos en el jardín. La abuela Ester está pintando la charca. Te enseña el cuadro: hay una ranita… ¡con un mechón rubio!", "dato": "El ciclo vuelve a empezar: huevo, renacuajo, rana… ¡y nuevos huevos! A esto se le llama ciclo de vida.", "sonido": "¡CROAC!", "obtiene": "", "imagen": "rana_23.webp"}, {"aventura_id": "rana", "num": 24, "tipo": "fin", "titulo": "Rana polo", "texto": "Te quedas en la orilla… y la escarcha de la mañana te deja tiesa como un polo. Tardas horas en volver a moverte.", "dato": "Al ser de sangre fría, con mucho frío las ranas apenas pueden moverse.", "sonido": "¡BRRR!", "obtiene": "", "imagen": "rana_24.webp"}, {"aventura_id": "rana", "num": 25, "tipo": "final", "titulo": "Reina de los mosquitos", "texto": "Pasas el verano cazando mosquitos con tu lengua. Las vecinas de la charca te aplauden… Y te despiertas en el jardín: ¡no te ha picado ni un mosquito! Papá, en cambio, está lleno de picaduras.", "dato": "Las ranas comen muchísimos insectos, como mosquitos: ayudan a mantener el equilibrio de la naturaleza.", "sonido": "¡HURRA!", "obtiene": "", "imagen": "rana_25.webp"}];
const RANA_QUIZ = [["rana", 0, 1, "¿Qué es un renacuajo?", "Una rana cuando es pequeña", "Un pez de agua dulce", "Un insecto"], ["rana", 0, 1, "¿Dónde ponen sus huevos las ranas?", "En el agua", "En un nido de ramas", "Bajo tierra"], ["rana", 1, 1, "¿Cómo respira un renacuajo?", "Con branquias", "Con pulmones", "No respira"], ["rana", 3, 1, "¿Qué comen sobre todo los renacuajos?", "Algas y restos de plantas", "Moscas", "Peces"], ["rana", 13, 1, "¿Con qué caza insectos una rana?", "Con su lengua pegajosa", "Con sus patas delanteras", "Con los dientes"], ["rana", 7, 2, "¿Qué patas le salen primero al renacuajo?", "Las traseras", "Las delanteras", "Todas a la vez"], ["rana", 0, 2, "¿Cómo se llama el cambio de renacuajo a rana?", "Metamorfosis", "Hibernación", "Migración"], ["rana", 10, 2, "¿A qué grupo de animales pertenece la rana?", "Anfibios", "Reptiles", "Peces"], ["rana", 5, 2, "¿Qué animal acuático caza renacuajos?", "La ninfa de libélula", "La mariposa", "La lombriz"], ["rana", 20, 2, "¿Dónde pasa el invierno la rana común?", "Enterrada en el barro del fondo", "En un nido de plumas", "Viaja a África"], ["rana", 0, 2, "¿Las ranas son vertebrados o invertebrados?", "Vertebrados: tienen esqueleto con columna", "Invertebrados", "Ninguna de las dos"], ["rana", 13, 3, "¿Qué le pasa a la cola del renacuajo?", "El cuerpo la absorbe poco a poco", "Se le cae de golpe", "Se convierte en una pata"], ["rana", 19, 3, "¿Quién croa en la rana común?", "Solo los machos", "Solo las hembras", "Los renacuajos"], ["rana", 20, 3, "¿Qué significa que la rana es de sangre fría?", "Su cuerpo tiene la temperatura del lugar donde está", "Su sangre es azul", "Siempre tiene frío"], ["rana", 14, 3, "¿Por qué los renacuajos tienen prisa por convertirse en ranas?", "Porque en verano muchas charcas se secan", "Porque se aburren", "Porque el agua sabe mal"], ["rana", 10, 3, "¿En qué se diferencia un pez de una rana?", "El pez respira por branquias toda su vida", "El pez tiene patas", "La rana vive siempre en el agua"]];

function _completarV12() {
  if (!_readSheet('Aventuras').some(a => String(a.id) === 'rana')) _append('Aventuras', RANA_AVENTURA);
  if (!_readSheet('Escenas').some(e => String(e.aventura_id) === 'rana')) RANA_ESCENAS.forEach(e => _append('Escenas', e));
  if (!_readSheet('Objetos').some(o => String(o.aventura_id) === 'rana')) RANA_OBJETOS.forEach(o => _append('Objetos', o));
  if (!_readSheet('Quiz').some(q => String(q.aventura_id) === 'rana')) RANA_QUIZ.forEach(q => _append('Quiz', {
    aventura_id: q[0], escena: q[1], nivel: q[2], pregunta: q[3], correcta: q[4], incorrecta_1: q[5], incorrecta_2: q[6] }));
}

/* ---------- v0.13: la familia entra en «De renacuajo a rana» ---------- */
function _completarV13() {
  const ss = _ss();
  const sync = (tab, keyCol, match, src, cols) => {
    const sh = ss.getSheetByName(tab), vals = sh.getDataRange().getValues(), H = vals[0].map(String);
    for (let r = 1; r < vals.length; r++) {
      const fila = src(vals[r], H); if (!fila) continue;
      cols.forEach(k => { const c = H.indexOf(k); if (c >= 0 && fila[k] != null && String(vals[r][c]) !== String(fila[k])) sh.getRange(r + 1, c + 1).setValue(fila[k]); });
    }
  };
  sync('Aventuras', 'id', null, (v, H) => String(v[H.indexOf('id')]) === 'rana' ? RANA_AVENTURA : null, ['intro']);
  sync('Escenas', 'aventura_id', null, (v, H) => String(v[H.indexOf('aventura_id')]) === 'rana' ? RANA_ESCENAS.find(e => e.num === Number(v[H.indexOf('num')])) : null, ['titulo', 'texto']);
}

/* ---------- v0.14: voz de Chewie con Google Cloud Text-to-Speech ----------
   La clave va en Configuración del proyecto › Propiedades del script › TTS_KEY (nunca en este código: el repo es público). */
const VOZ_CHEWIE = {
  name: 'Algieba', languageCode: 'es-419', model: 'gemini-3.1-flash-tts-preview', speakingRate: 1,
  prompt: 'Read aloud in a warm, welcoming and playful tone, like a curious little male hamster.'
};

function _tts(texto, voz) {
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('TTS_KEY');
  // Las voces Gemini no aceptan clave de API: van con tu propia sesión de Google (permiso cloud-platform) y tu proyecto.
  const oauth = !!voz.model, proyecto = props.getProperty('GCP_PROJECT') || '21070261983';
  if (!oauth && !key) throw new Error('Falta TTS_KEY en Propiedades del script.');
  const audioConfig = { audioEncoding: 'MP3', speakingRate: voz.speakingRate || 1 };
  if (voz.pitch != null && !voz.model && !/Chirp|Journey/i.test(voz.name)) audioConfig.pitch = voz.pitch;   // esas voces no admiten tono
  const input = { text: texto }; if (voz.prompt) input.prompt = voz.prompt;
  const voice = { languageCode: voz.languageCode || voz.name.split('-').slice(0, 2).join('-'), name: voz.name }; if (voz.model) voice.modelName = voz.model;
  const r = UrlFetchApp.fetch('https://texttospeech.googleapis.com/v1/text:synthesize' + (oauth ? '' : '?key=' + encodeURIComponent(key)), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: oauth ? { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'x-goog-user-project': proyecto } : {},
    payload: JSON.stringify({ input: input, voice: voice, audioConfig: audioConfig }) });
  const j = JSON.parse(r.getContentText());
  if (!j.audioContent) throw new Error('Google TTS: ' + r.getContentText().slice(0, 300));
  return Utilities.base64Decode(j.audioContent);
}
function _carpeta(nombre, padre) {
  const it = padre ? padre.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : (padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre));
}

/** 1) Genera muestras con las voces en español disponibles (normal y aguda) en Drive › AventurApp voces › pruebas. */
function probarVoces() {
  const key = PropertiesService.getScriptProperties().getProperty('TTS_KEY');
  const r = JSON.parse(UrlFetchApp.fetch('https://texttospeech.googleapis.com/v1/voices?languageCode=es&key=' + encodeURIComponent(key), { muteHttpExceptions: true }).getContentText());
  if (!r.voices) throw new Error(JSON.stringify(r).slice(0, 300));
  const voces = r.voices.map(v => v.name).filter(n => /^es-(ES|US)-(Neural2|Wavenet|Studio|Chirp3-HD)/.test(n)).sort();
  const f = _carpeta('pruebas ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd-MM HH.mm'), _carpeta('AventurApp voces'));
  const texto = '¡Psst! A los renacuajos les salen primero las patas traseras. Las delanteras salen después… ¡y la cola desaparece poco a poco!';
  let n = 0;
  voces.forEach(name => {
    const chirp = /Chirp/.test(name);
    (chirp ? [0] : [0, 8]).forEach(pitch => {
      if (n >= 40) return;
      try { f.createFile(Utilities.newBlob(_tts(texto, { name: name, pitch: pitch, speakingRate: 1.1 }), 'audio/mpeg', name + (pitch ? ' aguda' : '') + '.mp3')); n++; }
      catch (e) { Logger.log('✗ ' + name + ': ' + e.message); }
    });
  });
  Logger.log('✅ ' + n + ' muestras. Escúchalas aquí: ' + f.getUrl());
}

/** 1b) Una sola prueba con la voz elegida (VOZ_CHEWIE) antes de generarlas todas. */
function probarChewie() {
  const f = _carpeta('AventurApp voces');
  const file = f.createFile(Utilities.newBlob(_tts('¡Psst! ¿Sabías que a los renacuajos les salen primero las patas traseras? Las delanteras salen después… ¡y la cola desaparece poco a poco!', VOZ_CHEWIE), 'audio/mpeg', 'prueba Chewie.mp3'));
  Logger.log('✅ Escúchala: ' + file.getUrl());
}

/** 2) Genera un MP3 por cada secreto de Chewie en Drive › AventurApp voces › audio. Salta los que ya existen;
    si se corta por tiempo, vuelve a ejecutarlo. Para rehacer uno, borra su archivo. */
function generarVoces() {
  if (!VOZ_CHEWIE.name) throw new Error('Primero elige una voz en VOZ_CHEWIE.name');
  const t0 = Date.now(), dir = _carpeta('audio', _carpeta('AventurApp voces'));
  const hay = {}; const it = dir.getFiles(); while (it.hasNext()) hay[it.next().getName()] = true;
  const pend = _readSheet('Escenas').filter(e => String(e.dato || '').trim()).map(e => ({
    file: e.aventura_id + '_' + String(e.num).padStart(2, '0') + '.mp3', dato: String(e.dato).trim() })).filter(e => !hay[e.file]);
  let n = 0;
  for (const e of pend) {
    if (Date.now() - t0 > 300000) break;
    dir.createFile(Utilities.newBlob(_tts(e.dato, VOZ_CHEWIE), 'audio/mpeg', e.file)); n++;
  }
  Logger.log((n === pend.length ? '✅ Listo: ' : '⏳ Faltan ' + (pend.length - n) + ', ejecútalo otra vez. Hechos: ') + n + '. Carpeta: ' + dir.getUrl());
}

/** Revisa que cada aventura se pueda jugar: destinos que existen, decisiones con 2 opciones,
    finales alcanzables. Solo informa, no cambia nada. */
function diagnosticar() {
  const esc = _readSheet('Escenas');
  _readSheet('Aventuras').forEach(a => {
    const E = {};
    esc.filter(e => String(e.aventura_id) === String(a.id)).forEach(e => { E[Number(e.num)] = e; });
    const problemas = [];
    if (!Object.keys(E).length) { Logger.log('· ' + a.id + ' — sin escenas (' + (a.activa === 'pronto' ? 'próximamente' : 'vacía') + ')'); return; }
    Object.keys(E).forEach(n => {
      const e = E[n], tipo = String(e.tipo || 'decision').toLowerCase();
      if (tipo === 'decision') {
        ['a', 'b', 'c'].forEach(k => {
          if (!e['opcion_' + k]) { if (k !== 'c') problemas.push('Escena ' + n + ': falta opcion_' + k); return; }
          if (!E[Number(e['destino_' + k])]) problemas.push('Escena ' + n + ': destino_' + k + ' (' + e['destino_' + k] + ') no existe');
          if (k !== 'c' && e['requiere_' + k] && !E[Number(e['sino_' + k])]) problemas.push('Escena ' + n + ': requiere_' + k + ' sin un sino_' + k + ' que exista');
          if (k === 'c' && !e.requiere_c) problemas.push('Escena ' + n + ': la opción C necesita requiere_c');
        });
      } else if (tipo !== 'fin' && tipo !== 'final') problemas.push('Escena ' + n + ': tipo "' + e.tipo + '" no vale (decision / fin / final)');
    });
    /* alcanzables desde la 1 */
    const vis = {}, q = [1];
    while (q.length) { const n = q.shift(); if (vis[n] || !E[n]) continue; vis[n] = true; ['destino_a', 'destino_b', 'destino_c', 'sino_a', 'sino_b'].forEach(f => { if (E[n][f]) q.push(Number(E[n][f])); }); }
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
