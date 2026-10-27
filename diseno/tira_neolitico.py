import base64, io, segno, html
from PIL import Image
S = '/tmp/claude-0/-home-claude/0d6a1482-0436-5ae8-9eb8-d3f6d871d71c/scratchpad/'
IMG = '/home/claude/av6/img/'
def b64(path): return base64.b64encode(open(path, 'rb').read()).decode()
def foto(n):
    im = Image.open(f'{IMG}escenas/neolitico_{n:02d}.webp').convert('RGB')
    w, h = im.size; nh = int(w / 2)
    im = im.crop((0, (h - nh) // 2, w, (h - nh) // 2 + nh)).resize((520, 260), Image.LANCZOS)
    b = io.BytesIO(); im.save(b, 'JPEG', quality=86); return base64.b64encode(b.getvalue()).decode()
def font(p): return b64(S + 'fonts/' + p)
F = {k: font(v) for k, v in {
  'fd6': 'fontsource-fredoka-5.0.0/files/fredoka-latin-600-normal.woff2',
  'fd7': 'fontsource-fredoka-5.0.0/files/fredoka-latin-700-normal.woff2',
  'nu7': 'fontsource-nunito-5.0.0/files/nunito-latin-700-normal.woff2',
  'nu8': 'fontsource-nunito-5.0.0/files/nunito-latin-800-normal.woff2'}.items()}

URL = 'https://nnbarboza.github.io/aventurapp/?qr=1'
qr = segno.make(URL, error='m')
b = io.BytesIO(); qr.save(b, kind='svg', scale=4, border=1, dark='#2B211B'); QR = base64.b64encode(b.getvalue()).decode()

# num, tipo, titulo, texto, (opA, dA), (opB, dB), dato, img
P = [
 (1,'d','Primavera en La Draga','Vives en una casa de madera sobre pilotes, junto al lago. Es primavera: ¿con quién vas?',('Con papá, a sembrar',2),('Con el abuelo y las cabras',3),'',1),
 (2,'d','¡A sembrar!','Papá te da semillas de trigo… pero tú llevas unas bellotas que encontraste ayer.',('Sembrar el trigo',6),('Sembrar tus bellotas',4),'',2),
 (3,'d','Las cabras','La cabra más traviesa sale corriendo hacia el lago.',('Correr detrás gritando',5),('Llamarla bajito con hojas',6),'',3),
 (4,'f','Un bosque de robles','En otoño… ¡salen robles bebés en vez de trigo! Tardarán cien años en crecer.',None,None,'sembraban trigo y cebada guardando semillas de la cosecha anterior.',5),
 (5,'f','¡Excursión de cabras!','Tus gritos asustan a TODAS las cabras. Os pasáis la tarde buscándolas.',None,None,'la cabra y la oveja llegaron con los primeros agricultores.',6),
 (6,'d','Mediodía','Mamá cocina gachas en una olla de cerámica. ¡Qué rico! ¿Y después?',('Cerámica con la abuela',7),('Al bosque con papá',8),'',10),
 (7,'d','Barro y conchas','La abuela te deja decorar una vasija de barro. ¿Con qué?',('Con las pezuñas de una cabra',9),('Con el borde de una concha',10),'',13),
 (8,'d','El tejo','Papá busca madera de tejo para hacer un arco. ¡Hay uno enorme!',('Trepar hasta arriba',11),('Cortar una rama baja',12),'',14),
 (9,'f','La cabra alfarera','La cabra pisotea TODAS las vasijas de la abuela. ¡Clac, clac, clac!',None,None,'decoraban la cerámica con el borde de una concha: se llama cardial.',16),
 (10,'d','¡Al fuego!','Tu vasija ha quedado preciosa. Ahora hay que cocerla para que se endurezca.',('Secarla antes al sol',12),('¡Al fuego ya!',13),'',15),
 (11,'f','Atascada en el tejo','No sabes bajar. Papá tarda toda la tarde en rescatarte.',None,None,'en La Draga se encontraron arcos de tejo de los más antiguos de Europa.',17),
 (12,'d','¡La cosecha!','Llega el verano y el trigo está dorado. Te dan una hoz con dientes de sílex.',('Segar a toda velocidad',14),('Segar con cuidado',15),'',23),
 (13,'f','¡Pum!','La vasija, aún húmeda, explota en el fuego. ¡Llueve barro sobre el abuelo!',None,None,'secaban el barro antes de cocerlo para que no estallara.',20),
 (14,'f','La coleta del abuelo','Segas trigo, hierba… ¡y la coleta del abuelo!',None,None,'en la cosecha ayudaba todo el poblado, también los niños.',27),
 (15,'ff','La gran cosecha','Guardáis las mejores semillas para la primavera y todo el poblado lo celebra. ¡Eres una agricultora de La Draga!',None,None,'La Draga tiene unos 7.300 años: uno de los primeros poblados de agricultores de la Península.',32),
]
e = html.escape
def op(letra, o):
    return f'<div class="op"><span class="l">{letra}</span><span class="t">{e(o[0])}</span><span class="d">→{o[1]}</span></div>'
def panel(p):
    n, t, tit, txt, a, b_, dato, im = p
    cls = {'d': '', 'f': 'fin', 'ff': 'feliz'}[t]
    h = f'<section class="pan {cls}"><header><span class="num">{n}</span><h2>{e(tit)}</h2></header>'
    h += f'<img class="foto" src="data:image/jpeg;base64,{foto(im)}" alt="">'
    h += f'<p class="txt">{e(txt)}</p>'
    if t == 'd':
        h += '<div class="ops">' + op('A', a) + op('B', b_) + '</div>'
    elif t == 'f':
        h += f'<p class="dato"><b>Dato real:</b> {e(dato)}</p><div class="pie"><span class="sello">FIN</span><span>Vuelve al 1 ↺</span></div>'
    else:
        h += f'<div class="pie"><span class="sello feliz">¡FINAL FELIZ!</span><span>Sigue al 16 →</span></div>'
    return h + '</section>'
qrpanel = f'''<section class="pan qr"><header><span class="num">16</span><h2>¿Quieres más?</h2></header>
<p class="txt" style="text-align:center;margin-top:2px">¿Quieres jugar la <b>aventura completa</b>?</p>
<img class="qrimg" src="data:image/svg+xml;base64,{QR}" alt="QR">
<p class="mini">33 escenas, 16 finales, medallas, el quiz de Chewie… ¡y el Paleolítico!</p>
<p class="scan">📱 Escanéame con la cámara del móvil</p></section>'''
tiras = []
for i in range(4):
    ps = P[i*4:(i+1)*4]
    body = ''.join(panel(p) for p in ps) + (qrpanel if i == 3 else '')
    head = 'TIRA 1 · ¡EMPIEZA AQUÍ!' if i == 0 else f'TIRA {i+1} · pega bajo la tira {i}'
    tiras.append(f'<div class="tira"><div class="tab {"t1" if i == 0 else ""}">{head}</div>{body}</div>')
logo = b64(IMG + 'marca/logo_cartel.webp')
doc = f'''<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Tira Neolítico</title><style>
@font-face{{font-family:Fd;font-weight:600;src:url(data:font/woff2;base64,{F['fd6']})}}
@font-face{{font-family:Fd;font-weight:700;src:url(data:font/woff2;base64,{F['fd7']})}}
@font-face{{font-family:Nu;font-weight:700;src:url(data:font/woff2;base64,{F['nu7']})}}
@font-face{{font-family:Nu;font-weight:800;src:url(data:font/woff2;base64,{F['nu8']})}}
@page{{size:A4;margin:0}}
*{{box-sizing:border-box}}
html,body{{margin:0;background:#fff}}
.pag{{width:210mm;height:297mm;padding:7mm 7mm 6mm;display:flex;flex-direction:column;gap:3mm;font-family:Nu,sans-serif;font-weight:700;color:#2B211B}}
.cab{{display:flex;align-items:center;gap:4mm;height:20mm}}
.cab img{{height:20mm}}
.cab h1{{font-family:Fd;font-weight:700;font-size:21pt;margin:0;line-height:1}}
.cab p{{margin:1mm 0 0;font-size:8.5pt;line-height:1.3;color:#4A3D35}}
.tiras{{flex:1;display:grid;grid-template-columns:repeat(4,1fr);gap:2.5mm;min-height:0}}
.tira{{border:1.4px dashed #2B211B;display:flex;flex-direction:column;min-height:0}}
.tab{{height:5mm;flex:none;font-family:Fd;font-weight:600;font-size:7.5pt;letter-spacing:.04em;display:flex;align-items:center;justify-content:center;background:repeating-linear-gradient(45deg,#fff 0 3px,#EADBC6 3px 5px)}}
.tab.t1{{background:#2B211B;color:#fff}}
.pan{{flex:1;min-height:0;overflow:hidden;padding:1.6mm 1.8mm;display:flex;flex-direction:column;gap:1.1mm;border-top:1.2px solid #2B211B}}
.tab+.pan{{border-top:0}}
.pan header{{display:flex;align-items:center;gap:1.6mm}}
.num{{width:6.4mm;height:6.4mm;flex:none;border:1.6px solid #2B211B;border-radius:50%;display:flex;align-items:center;justify-content:center;font-family:Fd;font-weight:700;font-size:10pt}}
.fin .num{{background:#2B211B;color:#fff}} .feliz .num,.qr .num{{background:#D9653B;border-color:#D9653B;color:#fff}}
h2{{font-family:Fd;font-weight:600;font-size:9.6pt;margin:0;line-height:1.05;text-transform:uppercase}}
.foto{{width:100%;aspect-ratio:2/1;object-fit:cover;border-radius:2mm;border:1px solid #2B211B;flex:none}}
.txt{{margin:0;font-size:7.6pt;line-height:1.28}}
.ops{{margin-top:auto;display:flex;flex-direction:column;gap:1mm}}
.op{{display:flex;align-items:center;gap:1.4mm;border:1.2px solid #2B211B;border-radius:2mm;padding:.9mm 1.4mm .9mm 1mm;min-height:7.2mm}}
.op .l{{width:4.4mm;height:4.4mm;flex:none;border:1.2px solid #2B211B;border-radius:50%;display:flex;align-items:center;justify-content:center;font-family:Fd;font-weight:700;font-size:7pt}}
.op .t{{flex:1;font-size:7.2pt;line-height:1.15}}
.op .d{{font-family:Fd;font-weight:700;font-size:10pt;color:#D9653B;white-space:nowrap}}
.fin{{background:#FFF4EC}} .feliz{{background:#FFF8DF}}
.dato{{margin:0;font-size:6.9pt;line-height:1.25;color:#4A3D35}} .dato b{{color:#D9653B}}
.pie{{margin-top:auto;display:flex;align-items:center;justify-content:space-between;font-size:7.4pt;font-weight:800}}
.sello{{font-family:Fd;font-weight:700;font-size:10pt;color:#D9653B;border:1.6px solid #D9653B;border-radius:1.5mm;padding:0 2mm;transform:rotate(-5deg)}}
.sello.feliz{{background:#D9653B;color:#fff;font-size:8pt}}
.qr{{background:#FFF8DF;align-items:center}} .qr header{{align-self:stretch}}
.qrimg{{width:27mm;height:27mm;margin:.5mm 0}}
.mini{{margin:0;font-size:6.8pt;line-height:1.25;text-align:center;color:#4A3D35}}
.scan{{margin:auto 0 0;font-family:Fd;font-weight:600;font-size:7.6pt;text-align:center}}
</style></head><body><div class="pag">
<div class="cab"><img src="data:image/webp;base64,{logo}" alt=""><div><h1>¿Sobrevivirás al Neolítico?</h1>
<p>✂ Recorta las 4 tiras, pégalas en orden y enróllalas. Empieza en el <b>1</b>, elige <b>A</b> o <b>B</b> y desenrolla hasta el número que te toque.<br>Hace 7.300 años, en La Draga (Banyoles, Girona): uno de los primeros poblados de agricultores.</p></div></div>
<div class="tiras">{"".join(tiras)}</div></div></body></html>'''
open(S + 'tira_neolitico.html', 'w').write(doc)
print('ok', len(doc) // 1024, 'KB')
