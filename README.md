# ESCOGE TU PROPIA AVENTURApp

Libro de «elige tu propia aventura» para el móvil. Cada decisión cambia la historia:
algunas acaban en un final desastroso (y gracioso), otras llevan al final feliz.
Todo lo que pasa es históricamente correcto.

- **Frontend**: un solo `index.html` (HTML + CSS + JS inline). Sin build ni dependencias.
- **Backend**: Google Apps Script (`Code.gs`) sobre un Google Sheet.
- **Hosting**: GitHub Pages. PWA instalable. Coste: 0 €.

## Qué tiene (v0.3.0)

| Sección | Qué hace |
|---|---|
| **¿Quién eres?** | Al abrir, cada niña elige su dibujo. El progreso es de cada una. |
| **Aventuras** | Tarjeta por aventura con su progreso (finales encontrados). Jugar / Seguir. |
| **Juego** | La historia se *desenrolla* como la tira de papel: cada elección añade una escena debajo. «Me arrepiento» deshace la última elección. Finales con sello FIN, onomatopeya, dato real y confeti en el final feliz. |
| **Finales** | Álbum: los finales encontrados con su dibujo; los que faltan salen como «Final secreto». |
| **Datos** | Cuaderno con los «¿Sabías que…?» de cada escena visitada. |
| **Mapa** | Árbol de escenas: lo visitado con número, lo que falta como **?**, el último camino en naranja. |

**Chewie** (el hámster) es el guía: cuenta el «¿Sabías que…?» de cada escena, consuela en los finales desastre y celebra el feliz. Viaja con ella por todas las épocas.

Funciona **sin backend** (modo local, con la aventura del Paleolítico incluida) — sirve para probarla ya.

## Puesta en marcha

1. **Repo**: crea `nnbarboza/aventurapp`, sube `index.html`, `manifest.json`, `sw.js` y la carpeta `img/`. Activa GitHub Pages en `main` (raíz). Ya se puede jugar en modo local.
2. **Sheet**: crea un Google Sheet «AVENTURApp - Datos» → Extensiones › Apps Script → pega `Code.gs` entero → Guardar.
3. En el editor, ejecuta **`setup()`** (acepta los permisos). Crea las pestañas y siembra el Paleolítico.
4. **Implementar › Nueva implementación › Aplicación web** · Ejecutar como: yo · Acceso: cualquier usuario. Copia la URL `/exec`.
5. En `index.html`, pega esa URL en `const BACKEND_URL = '';` y súbelo.
6. Jugadoras: pestaña **Usuarios** (`setup()` ya crea a Gina).

Cambios en `Code.gs`: pegar entero → Guardar → **Gestionar implementaciones › ✏️ › Versión nueva**. Guardar no basta.
Al subir versión: `APP_VERSION` (index.html), `VERSION` (Code.gs) y `CACHE` (sw.js) a la vez.

## El Sheet

| Pestaña | Columnas |
|---|---|
| **Aventuras** | `id · titulo · epoca · intro · color · emoji · dibujo · orden · activa · kit · fondo · prota · elementos · imagen` |
| **Escenas** | `aventura_id · num · tipo · titulo · texto · dato · opcion_a · destino_a · opcion_b · destino_b · dibujo · emoji · sonido · fondo · prota · elementos · imagen` |
| **Usuarios** | `username · nombre · avatar · color · activo` (avatar = ruta dentro de `img/`, ej. `kit/paleolitico/gina/contenta.webp`) |
| **Partidas** | `id · ts · aventura_id · recorrido · final · jugadora` (la escribe la app: una fila por final alcanzado) |
| **Config** | `clave · valor · nota` |

Se lee por **nombre de cabecera**: puedes reordenar columnas o añadir otras sin romper nada.
Los finales encontrados, los datos y el mapa **se calculan** desde `Partidas`; no se guardan.

### Cómo escribir una aventura nueva

1. Una fila en **Aventuras** con un `id` corto sin espacios (`egipto`, `romanos`…). `activa` = `TRUE`.
2. En **Escenas**, una fila por escena, todas con ese `aventura_id`. La historia empieza en `num` = 1.
3. `tipo`:
   - `decision` → lleva `opcion_a` + `destino_a` y `opcion_b` + `destino_b` (el número de escena al que lleva).
   - `fin` → final prematuro (el desastre gracioso). Sin opciones.
   - `final` → final feliz. Sin opciones.
4. `dato`: el «¿Sabías que…?» de esa escena. En los finales sale como «Dato real».
5. `sonido`: onomatopeya opcional para los finales (`¡PLOF!`).
6. Dos escenas pueden llevar al mismo final: así la historia no es un árbol plano.
7. Ejecuta **`diagnosticar()`** en el editor: avisa de destinos que no existen, escenas a las que no se puede llegar o si falta el final feliz.

Recomendación de forma (como el Paleolítico): ~16 escenas, 6 finales prematuros, 1 final feliz, 2 caminos buenos, y que la opción correcta vaya alternando entre A y B.

## Imágenes de las escenas (lo principal)

Cada escena muestra **su imagen** (columna `imagen` del Sheet), que se genera con ChatGPT en el estilo del kit y se sube a `img/escenas/`.
En `diseno/IMAGENES_PALEOLITICO.md` están el prompt base y la descripción de cada imagen con su nombre de fichero.

- Formato **3:2** (1536 × 1024), `.webp`, nombre en minúscula.
- Son las de **Gina** (`usuario_base` en Config). Otra jugadora puede tener las suyas en `img/escenas/<usuario>/` con los mismos nombres; las que falten se ven con las de Gina.
- Portada de cada aventura: columna `imagen` de Aventuras → `img/aventuras/`.
- Mientras una imagen no exista, la escena se compone con las piezas del kit (abajo). Nada se rompe.

## Plan B: escenas compuestas con el kit

Cada escena se **compone** con piezas del UI kit, desde el Sheet:

| Columna | Qué pones | Ejemplo |
|---|---|---|
| `fondo` | un escenario de la época | `cueva_dentro` |
| `prota` | la expresión de la protagonista | `miedo` |
| `elementos` | hasta 3 piezas separadas por comas (animales, familia, objetos, efectos) | `oso,zzz` |
| `imagen` | *opcional*: una ilustración completa que sustituye a todo lo anterior | `paleolitico_08.webp` → `img/escenas/paleolitico_08.webp` |

En `Aventuras`, la columna `kit` dice qué juego de piezas usa la época (`paleolitico`). Una aventura nueva de otra época = una carpeta nueva `img/kit/<epoca>/` con sus piezas, y la prota con su ropa de esa época.

```
img/kit/
├── paleolitico/
│   ├── gina/       normal contenta entusiasmada pensando idea sorpresa miedo enfadada triste cansada
│   │               corriendo celebrando enamorada llorando mareada asombrada asustada cantando asco agotada
│   ├── clan/       papa mama hermano abuela abuelo
│   ├── animales/   mamut reno bisonte oso lobo jabali ciervo liebre ave pez  (+ leon, pendiente)
│   ├── objetos/    lanza propulsor cuchillo hueso rama antorcha bolsa carne bayas plantas silex piel
│   │               trampa fuego refugio cueva huellas piedras nieve
│   └── fondos/     campamento cueva bosque montanas rio valle nieve cueva_dentro
└── comun/          (sirve en todas las épocas)
    ├── chewie/     normal oliendo comiendo corriendo dormido sorprendido sentado explorando
    └── fx/         estrella exclamacion interrogacion musica zzz lineas nube copo gotas flecha
                    piedritas hierba huella corazon garabato
```

**Las piezas actuales están recortadas del kit** (`diseno/ui-kit.png`): funcionan, pero son pequeñas (~100 px) y se ven algo borrosas. Para mejorarlas, regenera cada una en alta (**~400 px de alto, `.webp` con fondo transparente**) y súbela **con el mismo nombre**: aparece sola.

Una pieza que no existe simplemente no se pinta (no rompe nada). Si una escena no tiene `fondo/prota/elementos`, usa el dibujo SVG (`dibujo`) o el `emoji`.

### Pendiente del kit (rigor histórico)
- `animales/leon.webp` **no está**: el del kit tiene melena y el león de las cavernas casi no tenía (así lo pintaron en Chauvet). Redibújalo sin melena y súbelo; la escena 13 ya lo pide.
- Las **raquetas de nieve** del kit no se usan: no hay pruebas de que existieran en el Paleolítico.
- Las gafas de sol de Gina, tampoco (sí, obvio 😄).

### Iconos de la app
`img/icon-192.png`, `icon-512.png`, `icon-maskable-512.png`, `apple-touch-icon.png` — provisionales (el mamut).

## Funciones de mantenimiento (editor de Apps Script)

| Función | Qué hace |
|---|---|
| `setup()` | Crea pestañas y columnas que falten; siembra el Paleolítico si no hay aventuras. Idempotente. |
| `diagnosticar()` | Revisa que cada aventura se pueda jugar. Solo informa. |
| `resembrarPaleoliticoAplicar()` | Reescribe el Paleolítico con la versión del código (si la editaste a mano, se pierde). |

## Rigor histórico (Paleolítico)

Ambientada hace ~15.000 años (Paleolítico superior, Magdalenense). Detalles cuidados a propósito:
propulsor (no arco), fuego con sílex y pirita, lámparas de grasa, agujas de hueso, flauta de hueso,
manos en negativo soplando ocre (El Castillo); la osa es **parda** (el oso de las cavernas ya se había
extinguido) y el león de las cavernas va **sin melena**, como en las pinturas de Chauvet.
