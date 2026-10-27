/* ============================================================
   AVENTURApp — backend (Google Apps Script)
   Se pega entero en el editor de Apps Script del Sheet.

   Primera vez: ejecutar setup() desde el editor (crea pestañas,
   guarda el ID del Sheet y siembra la aventura de ejemplo).
   Luego: Implementar › Nueva implementación › Aplicación web
          Ejecutar como: yo · Acceso: cualquier usuario.
   Cada cambio de código: Gestionar implementaciones › ✏️ › Versión nueva.
   ============================================================ */
const VERSION = '0.8.0';

const HEADERS = {
  Usuarios:  ['username', 'nombre', 'avatar', 'color', 'activo'],
  Aventuras: ['id', 'titulo', 'epoca', 'intro', 'color', 'emoji', 'dibujo', 'orden', 'activa',
              'kit', 'fondo', 'prota', 'elementos', 'imagen', 'hace'],
  Escenas:   ['aventura_id', 'num', 'tipo', 'titulo', 'texto', 'dato',
              'opcion_a', 'destino_a', 'opcion_b', 'destino_b', 'dibujo', 'emoji', 'sonido',
              'fondo', 'prota', 'elementos', 'imagen'],
  Partidas:  ['id', 'ts', 'aventura_id', 'recorrido', 'final', 'jugadora'],
  Quiz:      ['aventura_id', 'escena', 'pregunta', 'correcta', 'incorrecta_1', 'incorrecta_2'],
  Quizzes:   ['id', 'ts', 'aventura_id', 'aciertos', 'total', 'jugadora'],
  Invitados: ['username', 'nombre', 'avatar', 'creado', 'ultimo'],
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
    quizzes:   _readSheet('Quizzes')
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
  [23, "decision", "¡La cosecha!", "Es verano y el trigo está dorado. Todo el poblado sale a segar y a ti te dan una hoz de madera con dientes de sílex.", "Las hoces eran de madera, con trocitos de sílex afilado encajados como dientes.", "Segar a toda velocidad para ganar a tu hermano", 27, "Segar con cuidado, como te enseñan", 26, "🌾", ""],
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
