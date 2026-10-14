<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Nómina Domingo Celi</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
/* Layout: panel de control a la izquierda (pizarrón), hojas A4 en un escritorio a la derecha; solo las hojas se imprimen */
:root{
  --bg:#e9ece7; --panel:#fbfcfa; --ink:#1b2420; --muted:#5b6862; --line:#d3d9d4;
  --accent:#24584a; --accent-ink:#ffffff; --accent-soft:#dfeae4; --warn:#9a5b12; --warn-soft:#f6ead8;
  --font-ui:"Archivo", "Segoe UI", system-ui, sans-serif;
  --font-mono:"IBM Plex Mono", ui-monospace, Consolas, monospace;
  --font-sheet: Arial, Helvetica, sans-serif;
  color-scheme: light;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --bg:#111614; --panel:#1a211e; --ink:#e4ebe7; --muted:#93a29a; --line:#2d3833;
    --accent:#77bba2; --accent-ink:#0e1a15; --accent-soft:#1f3029; --warn:#e2a95c; --warn-soft:#2e2416;
    color-scheme: dark;
  }
}
:root[data-theme="dark"]{
  --bg:#111614; --panel:#1a211e; --ink:#e4ebe7; --muted:#93a29a; --line:#2d3833;
  --accent:#77bba2; --accent-ink:#0e1a15; --accent-soft:#1f3029; --warn:#e2a95c; --warn-soft:#2e2416;
  color-scheme: dark;
}
*{box-sizing:border-box}
html,body{margin:0}
body{background:var(--bg);color:var(--ink);font:15px/1.5 var(--font-ui)}
.app{display:grid;grid-template-columns:340px minmax(0,1fr);gap:24px;align-items:start;
  padding-block:24px;padding-left:max(16px, env(safe-area-inset-left, 0px));padding-right:max(16px, env(safe-area-inset-right, 0px));max-width:1320px;margin:0 auto}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:20px;display:flex;flex-direction:column;gap:18px;position:sticky;top:calc(16px + env(safe-area-inset-top, 0px));max-height:calc(100vh - 32px);overflow-y:auto}
.brand h1{font-size:20px;line-height:1.2;margin:0;font-weight:700;text-wrap:balance}
.brand p{margin:4px 0 0;color:var(--muted);font-size:13.5px}
.eyebrow{font:500 11px/1 var(--font-mono);letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}
.drop{border:1.5px dashed var(--line);border-radius:8px;padding:18px 14px;text-align:center;cursor:pointer;display:flex;flex-direction:column;gap:6px;align-items:center;transition:border-color .15s, background .15s}
.drop:hover,.drop.over{border-color:var(--accent);background:var(--accent-soft)}
.drop:focus-within{outline:2px solid var(--accent);outline-offset:2px}
.drop strong{font-weight:600}
.drop span{color:var(--muted);font-size:13px}
.drop input{position:absolute;width:1px;height:1px;opacity:0}
.status{font-size:13.5px;border-radius:6px;padding:9px 11px;background:var(--accent-soft);color:var(--ink)}
.status.warn{background:var(--warn-soft);color:var(--warn)}
.fields{display:flex;flex-direction:column;gap:12px}
.fields label{display:flex;flex-direction:column;gap:4px;font-size:13px;font-weight:500;color:var(--muted)}
.fields input,.fields select{font:15px var(--font-ui);color:var(--ink);background:var(--bg);border:1px solid var(--line);border-radius:6px;padding:8px 10px;width:100%}
.fields input:focus,.fields select:focus{outline:2px solid var(--accent);outline-offset:0;border-color:transparent}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.check{flex-direction:row!important;align-items:center;gap:8px!important;color:var(--ink)!important;font-weight:400!important}
.check input{width:auto}
.actions{display:flex;flex-direction:column;gap:8px}
button{font:600 15px var(--font-ui);border-radius:7px;padding:11px 14px;cursor:pointer;border:1px solid var(--line);background:transparent;color:var(--ink)}
button.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
button:hover{filter:brightness(1.05)}
button:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.hint{font-size:12.5px;color:var(--muted);margin:0}
.cfg{display:flex;flex-direction:column;gap:10px}
.stepper{display:grid;grid-template-columns:44px 1fr 44px;border:1px solid var(--line);border-radius:7px;overflow:hidden;background:var(--bg)}
.stepper button{border:0;border-radius:0;padding:8px 0;font-size:20px;line-height:1;background:transparent;color:var(--accent)}
.stepper button:hover{background:var(--accent-soft)}
.stepper input{border:0!important;border-left:1px solid var(--line)!important;border-right:1px solid var(--line)!important;border-radius:0!important;text-align:center;font:600 20px var(--font-ui)!important;font-variant-numeric:tabular-nums;-moz-appearance:textfield;appearance:textfield}
.stepper input::-webkit-inner-spin-button,.stepper input::-webkit-outer-spin-button{-webkit-appearance:none;margin:0}
.presets{display:flex;flex-wrap:wrap;gap:6px}
.presets button{font:500 13px var(--font-mono);padding:5px 11px;border-radius:999px;font-variant-numeric:tabular-nums}
.presets button[aria-pressed="true"]{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
.fields select:disabled{opacity:.6}
.seg{display:grid;grid-template-columns:1fr 1fr;border:1px solid var(--line);border-radius:7px;overflow:hidden}
.seg button{border:0;border-radius:0;font-size:14px;padding:10px 6px;color:var(--muted)}
.seg button+button{border-left:1px solid var(--line)}
.seg button[aria-pressed="true"]{background:var(--accent);color:var(--accent-ink)}
.empty{align-self:stretch;border:1.5px dashed var(--line);border-radius:10px;min-height:340px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;text-align:center;padding:32px 20px;color:var(--muted)}
.empty strong{font-size:18px;color:var(--ink)}
.empty span{max-width:46ch}
.empty .mono{font:500 12px var(--font-mono);letter-spacing:.04em;margin-top:6px}
button:disabled{opacity:.5;cursor:not-allowed}
[hidden]{display:none!important}
.groups{display:flex;flex-wrap:wrap;gap:6px}
.filt{display:flex;flex-direction:column;gap:6px}
.flabel{font-size:13px;font-weight:500;color:var(--muted)}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chips button{font:500 14px var(--font-ui);padding:6px 12px;border-radius:999px;display:inline-flex;gap:6px;align-items:baseline}
.chips button small{font:500 11px var(--font-mono);color:var(--muted)}
.chips button[aria-pressed="true"]{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
.chips button[aria-pressed="true"] small{color:inherit;opacity:.8}

.stage{min-width:0;overflow-x:auto;display:flex;flex-direction:column;align-items:center;gap:24px;padding-bottom:24px}
.stage-head{align-self:stretch;display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap}
.stage-head .eyebrow{font-size:11.5px}

/* ---------- Hoja A4 (siempre papel blanco, también en modo oscuro) ---------- */
.sheet{background:#fff;color:#000;width:210mm;min-height:297mm;padding:9mm 10mm 10mm;font-family:var(--font-sheet);font-size:8.6pt;
  box-shadow:0 1px 2px rgba(0,0,0,.12),0 8px 24px rgba(0,0,0,.10);flex:none;display:flex;flex-direction:column}
.sheet .top{display:grid;grid-template-columns:20mm 1fr auto;align-items:center;gap:4mm;margin-bottom:1.6mm}
.sheet .logo{width:20mm;height:20mm;object-fit:contain;display:block}
.sheet .inst{text-align:center;font-weight:700;font-size:11pt;padding-bottom:.6mm}
.sheet .anio{font-weight:700;font-size:9.5pt;white-space:nowrap;padding-bottom:.6mm}
.sheet table{border-collapse:collapse;table-layout:fixed;width:100%}
.sheet th,.sheet td{padding:0;font-weight:400}
.sheet .hbox{border:1px solid #000;text-align:left;vertical-align:top;padding:1.4mm 2mm 1.6mm}
.sheet .hsec{text-align:center;font-size:8.6pt}
.sheet .hgr{text-align:center;font-size:7.6pt;margin-bottom:3mm}
.sheet .fld{display:flex;align-items:flex-end;gap:2mm;margin-top:2.2mm;font-size:8.4pt}
.sheet .fld .lbl{white-space:nowrap}
.sheet .fld .ln{border-bottom:1px solid #000;min-height:3.6mm;line-height:1.1;padding:0 1mm;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sheet .ln.short{width:12mm;flex:none}
.sheet .ln.mid{flex:0 1 62mm;min-width:0}
.sheet .ln.long{flex:0 1 76mm;min-width:0}
.sheet .gh{border:1px solid #000}
.sheet tbody td.g{border:1px solid #000;height:4.75mm}
.sheet td.n{text-align:center;vertical-align:top;font-size:8pt;line-height:4.75mm}
.sheet td.nm{white-space:nowrap;overflow:hidden;padding-left:1mm;padding-right:1mm}
.sheet .nmw{display:flex;align-items:baseline;gap:1.6mm;line-height:4.75mm}
.sheet .cd{flex:none}
.sheet .name{flex:none;max-width:100%;overflow:hidden;text-overflow:ellipsis}
.sheet .lead{flex:1;min-width:4mm;overflow:hidden;white-space:nowrap;letter-spacing:.9mm;color:#000}
.sheet .sign{margin-top:16mm;display:flex;justify-content:flex-end;gap:8mm;align-items:flex-end;font-size:9pt;padding-right:2mm}
.sheet .sign .ln{display:inline-block;border-bottom:1px solid #000;width:58mm;min-height:4mm;padding:0 1mm;font-size:8.6pt}
.sheet .cont{font-size:7.5pt;text-align:right;margin-top:auto;padding-top:4mm;color:#333}

@media (max-width:900px){
  .app{grid-template-columns:minmax(0,1fr)}
  .panel{position:static}
  .stage{align-items:flex-start}
}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}

@media print{
  @page{size:A4 portrait;margin:0}
  html,body{background:#fff!important}
  .panel,.stage-head{display:none!important}
  .app{display:block;padding:0;max-width:none}
  .stage{display:block;padding:0;overflow:visible;gap:0}
  .sheet{box-shadow:none;width:210mm;height:297mm;min-height:0;overflow:hidden;break-after:page;page-break-after:always;margin:0}
  .sheet:last-child{break-after:auto;page-break-after:auto}
  *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
}
</style>
</head>
<body>
<div class="app">
  <aside class="panel" aria-label="Controles">
    <div class="brand">
      <div class="eyebrow">Registro docente</div>
      <h1>Nómina por curso</h1>
      <p>Sube el Excel de estudiantes y obtén la hoja lista para imprimir.</p>
    </div>

    <label class="drop" id="drop" for="file">
      <input type="file" id="file" accept=".xlsx,.xls,.xlsm,.csv,.ods">
      <strong>Subir archivo de Excel</strong>
      <span>Arrastra aquí o haz clic · .xlsx, .xls, .csv</span>
    </label>
    <div class="status" id="status">Sube el archivo de Excel del curso. Los estudiantes, la matrícula, el año lectivo, la sección, el grado y el paralelo se toman de ahí.</div>

    <div class="fields" id="filters" hidden>
      <div class="eyebrow">Filtrar</div>
      <div class="filt"><span class="flabel">Curso</span><div class="chips" id="fltCurso" role="group" aria-label="Filtrar por curso"></div></div>
      <div class="filt"><span class="flabel">Paralelo</span><div class="chips" id="fltPar" role="group" aria-label="Filtrar por paralelo"></div></div>
    </div>

    <div class="fields">
      <div class="eyebrow">Columnas del listado</div>
      <div class="cfg">
        <div class="stepper" role="group" aria-label="Número de columnas">
          <button type="button" id="colsMinus" aria-label="Quitar una columna">−</button>
          <input id="fCols" type="number" min="0" max="25" value="11" aria-label="Número de columnas">
          <button type="button" id="colsPlus" aria-label="Agregar una columna">+</button>
        </div>
        <div class="presets" id="presets">
          <button type="button" data-n="5">5</button><button type="button" data-n="8">8</button><button type="button" data-n="10">10</button><button type="button" data-n="11">11</button><button type="button" data-n="15">15</button><button type="button" data-n="20">20</button>
        </div>
        <p class="hint">Columnas en blanco para notas o asistencia, de 0 a 25. El ancho se ajusta solo a la hoja A4.</p>
      </div>
    </div>

    <div class="fields">
      <div class="eyebrow">Datos del Excel</div>
      <div class="seg" role="group" aria-label="Orden del nombre en la lista">
        <button type="button" id="ordApe" aria-pressed="true">Apellidos primero</button>
        <button type="button" id="ordNom" aria-pressed="false">Nombres primero</button>
      </div>
      <div class="row2">
        <label>Columnas de apellidos<input id="fApe" type="text" value="E, F" placeholder="E, F" autocomplete="off"></label>
        <label>Columnas de nombres<input id="fNom" type="text" value="G, H" placeholder="G, H" autocomplete="off"></label>
      </div>
      <p class="hint">Solo se leen estas columnas del Excel. Los datos de padres o representantes no se usan.</p>
      <label>Columna de matrícula<input id="fMat" type="text" value="AI" placeholder="AI" autocomplete="off"></label>
      <p class="hint">La matrícula de cada estudiante se lee siempre de esta columna, en la misma fila de sus apellidos y nombres.</p>
      <label class="check"><input id="fSort" type="checkbox"> Ordenar alfabéticamente</label>
    </div>

    <div class="fields">
      <div class="eyebrow">Encabezado</div>
      <label>Institución<input id="fInst" type="text"></label>
      <div class="row2">
        <label>Año lectivo<input id="fAnio" type="text" placeholder="Del Excel"></label>
        <label>Trimestre<input id="fTrim" type="text" placeholder="Ej.: 1"></label>
      </div>
      <label>Asignatura<input id="fAsig" type="text" placeholder="Ej.: Matemática"></label>
      <label>Profesor guía<input id="fGuia" type="text" placeholder="Nombre del tutor"></label>
      <label>Nombre del profesor (firma)<input id="fProf" type="text" placeholder="Opcional"></label>
    </div>

    <div class="actions">
      <button class="primary" id="btnPrint" type="button">Imprimir / Guardar PDF</button>
      <button id="btnTpl" type="button">Descargar plantilla de Excel</button>
      <p class="hint">Año lectivo, sección, grado y paralelo se leen del Excel. Si escribes el año lectivo aquí, reemplaza al del archivo.</p>
    </div>
  </aside>

  <main class="stage" id="stage" aria-label="Vista previa">
    <div class="stage-head"><span class="eyebrow" id="pageCount"></span></div>
    <div class="empty" id="empty">
      <strong>Aún no hay nómina</strong>
      <span>Sube el archivo de Excel en el panel de la izquierda. La hoja aparecerá aquí con los estudiantes del archivo, lista para imprimir.</span>
      <span class="mono">Apellidos en E y F · Nombres en G y H · Matrícula en su columna</span>
    </div>
  </main>
</div>

<script>
(function(){
"use strict";
const ROWS_PER_PAGE = 40;
const LOGO = "data:image/png;base64,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";
const DEFAULT_INST = "UNIDAD EDUCATIVA FISCOMISIONAL DOMINGO CELI";


let groups = [];

/* ---------- Utilidades ---------- */
const $ = id => document.getElementById(id);
const clean = v => String(v == null ? "" : v).replace(/\s+/g," ").trim();
const norm = v => clean(v).normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase();
const esc = s => String(s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const stripDots = s => clean(s).replace(/(\s*\.)+\s*$/,"").trim();
const CODE_RE = /^\d{4}-\d{2,6}$/;
const YEAR_RE = /(20\d{2})\s*[-–\/]\s*(20\d{2})/;

const ORD = {1:"Primer",2:"Segundo",3:"Tercer",4:"Cuarto",5:"Quinto",6:"Sexto",7:"Séptimo",8:"Octavo",9:"Noveno",10:"Décimo"};
const ORD_WORD = /^(primer|primero|segundo|tercer|tercero|cuarto|quinto|sexto|s[eé]ptimo|octavo|noveno|d[eé]cimo)$/i;
const cap = s => s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();

function fmtGrado(g){
  g = clean(g); if(!g) return "";
  let m = g.match(/^(\d{1,2})\s*(°|º|o|\.?o)?\s*(grado|egb)?\.?$/i);
  if(m && ORD[+m[1]]) return ORD[+m[1]] + " Grado de Educación General Básica";
  if(ORD_WORD.test(g)) return cap(g) + " Grado de Educación General Básica";
  m = g.match(/^(\S+)\s+(grado|egb)$/i);
  if(m && ORD_WORD.test(m[1])) return cap(m[1]) + " Grado de Educación General Básica";
  return g;
}
function fmtSec(s){
  s = clean(s); if(!s) return "";
  return /secci/i.test(s) ? s.toUpperCase() : "SECCIÓN " + s.toUpperCase();
}
function fmtPar(p){
  p = clean(p); if(!p) return "";
  return /paralelo/i.test(p) ? p.replace(/paralelo/i,"Paralelo") : "Paralelo " + p.toUpperCase();
}
function fmtAnio(a){
  a = clean(a); const m = a.match(YEAR_RE);
  return m ? m[1] + "-" + m[2] : a;
}

/* ---------- Lectura del Excel ---------- */
function classify(h){
  const n = norm(h); if(!n) return null;
  // Columnas de padres / representantes: se ignoran por completo
  if(/padre|madre|representante|tutor|familiar|conyuge|apoderado|acudiente|responsable|papa|mama/.test(n)) return null;
  const n2 = n.replace(/[°º.#\s]/g,"");
  if(n.length <= 8 && ["","n","no","nro","num","numero","orden","no."].includes(n2)) return "num";
  if(/apellidos?\s*(y|,)?\s*nombres?|nombres?\s*(y|,)?\s*apellidos?|estudiante|alumno|nombre completo|nomina|discente/.test(n)) return "full";
  if(/apellido|paterno|materno/.test(n)) return "ape";
  if(/nombre/.test(n) && !/representante|padre|madre|docente|profesor|tutor|institucion|usuario|curso|grado|archivo|hoja/.test(n)) return "nom";
  if(/matricula/.test(n) && !/fecha|estado|tipo|valor|costo/.test(n)) return "mat";
  if(/codigo|^cod\b/.test(n)) return "code";
  if(/cedula|identificacion|^ci$|^dni$|documento/.test(n)) return "ced";
  if(/ano lectivo|periodo lectivo|^periodo|^ano$/.test(n)) return "anio";
  if(/seccion|jornada/.test(n)) return "sec";
  if(/paralelo/.test(n)) return "par";
  if(/grado|curso|nivel|ano basico|ano de educacion/.test(n)) return "grado";
  if(/institucion|unidad educativa/.test(n)) return "inst";
  return null;
}

function scanMeta(rows, upto, meta){
  const LBL = [
    [/^a[ñn]o\s+lectivo|^per[ií]odo(\s+lectivo)?/i,"anio"],
    [/^secci[oó]n|^jornada/i,"sec"],
    [/^paralelo/i,"par"],
    [/^grado|^curso|^nivel/i,"grado"],
    [/^instituci[oó]n/i,"inst"]
  ];
  for(let r = 0; r < Math.min(upto, rows.length); r++){
    const row = rows[r] || [];
    for(let c = 0; c < row.length; c++){
      const t = clean(row[c]); if(!t) continue;
      let used = false;
      for(const [re,key] of LBL){
        if(!re.test(t)) continue;
        const after = t.replace(re,"").replace(/^\s*[:\-–]\s*/,"").trim();
        let val = after;
        if(!val){ for(let k = c+1; k < row.length; k++){ if(clean(row[k])){ val = clean(row[k]); break; } } }
        if(key === "sec" && /^(secci[oó]n)\s+(matutina|vespertina|nocturna)/i.test(t)) val = t;
        if(val && !meta[key]) { meta[key] = val; used = true; }
        break;
      }
      if(used) continue;
      // Texto libre
      if(!meta.inst && /unidad educativa|colegio|escuela|instituto/i.test(t) && t.length > 12) meta.inst = t.replace(/\s*a[ñn]o lectivo.*$/i,"").trim();
      if(!meta.anio && /a[ñn]o lectivo/i.test(t) && YEAR_RE.test(t)) meta.anio = t.match(YEAR_RE)[0];
      if(!meta.sec){ const m = t.match(/\b(matutina|vespertina|nocturna)\b/i); if(m) meta.sec = /secci/i.test(t) ? t : m[1]; }
      if(/grado|curso|bachillerato|b[aá]sica/i.test(t) && /educaci|grado|bachillerato/i.test(t) && !/unidad educativa/i.test(t)){
        const pm = t.match(/^(.*?)\s*paralelo\s*[:\-]?\s*([A-Z0-9]{1,3})\b/i);
        if(pm){ if(!meta.grado) meta.grado = pm[1]; if(!meta.par) meta.par = pm[2]; }
        else if(!meta.grado) meta.grado = t;
      } else if(!meta.par){
        const pm = t.match(/^paralelo\s*[:\-]?\s*([A-Z0-9]{1,3})$/i); if(pm) meta.par = pm[1];
      }
    }
  }
  return meta;
}

const SUBHDR = /^(apellidos?|nombres?|paterno|materno|primer[oa]?|segundo|1er|2do|1|2)(\s+(nombre|apellido)s?)?$/;
const isAlpha = v => /[A-Za-zÁÉÍÓÚÑÜáéíóúñü]/.test(v) && !/\d/.test(v);
const joinCols = (row, cols) => cols.map(c => stripDots(clean(row[c]))).filter(Boolean).join(" ");

// opts: {codeCol, apeCols, nomCols, manual}
function parseSheet(rows, sheetName, opts){
  opts = opts || {};
  // 1) Encontrar fila de encabezados
  let hRow = -1, map = {};
  for(let r = 0; r < Math.min(rows.length, 40); r++){
    const row = rows[r] || []; const m = {apes:[], noms:[]}; let hits = 0;
    row.forEach((cell,c) => {
      const k = classify(cell); if(!k) return;
      if(k === "ape" || k === "nom") m[k + "s"].push(c);
      if(m[k] == null){ m[k] = c; hits++; }
    });
    const colsHdr = [...(opts.apeCols || []), ...(opts.nomCols || [])].some(c => ["ape","nom","full"].includes(classify(row[c])) || SUBHDR.test(norm(clean(row[c]))));
    if(colsHdr || (hits >= 2 && (m.full != null || m.ape != null || m.nom != null || m.mat != null || m.code != null))){ hRow = r; map = m; break; }
  }
  const meta = scanMeta(rows, hRow >= 0 ? hRow : 15, {});
  if(!meta.grado && /grado|curso|°|º/i.test(sheetName||"")) meta.grado = sheetName;

  const students = [];
  const BAD = /^(total|nombre profesor|f\.?:?$|firma|profesor|docente|observaci)/i;
  let headers = [], codeLabel = "", apeCols = [], nomCols = [], fullCol = null;

  const pushStudent = (row, codeCol, extra) => {
    const ape = joinCols(row, apeCols), nom = joinCols(row, nomCols);
    const full = fullCol != null ? stripDots(clean(row[fullCol])) : "";
    const shown = full || [ape, nom].filter(Boolean).join(" ");
    if(!shown || /^\d+([.,]\d+)?$/.test(shown) || BAD.test(shown)) return;
    if(!full && [...apeCols, ...nomCols].some(c => classify(row[c]) || SUBHDR.test(norm(clean(row[c]))))) return;   // filas de encabezado
    students.push(Object.assign({
      code: codeCol != null ? clean(row[codeCol]) : "",
      ape: ape.toUpperCase(), nom: nom.toUpperCase(), full: full.toUpperCase(),
      anio:"", sec:"", grado:"", par:"", inst:""
    }, extra || {}));
  };

  if(hRow >= 0){
    const hdr = (rows[hRow] || []).map(clean);
    headers = hdr.filter(Boolean);
    const data = rows.slice(hRow + 1);
    // Apellidos / nombres: columnas elegidas por el usuario o detectadas
    if(opts.manual){ apeCols = opts.apeCols || []; nomCols = opts.nomCols || []; }
    else {
      apeCols = map.apes.slice(); nomCols = map.noms.slice();
      // Encabezados combinados (p. ej. "APELLIDOS" sobre E:F): sumar columnas vecinas sin título que traen texto
      const taken = new Set(Object.values(map).filter(v => typeof v === "number"));
      const extend = cols => { if(cols.length !== 1) return cols; const out = cols.slice(); let k = cols[0] + 1;
        while(k < hdr.length + 3 && !hdr[k] && !taken.has(k)){
          const vals = data.map(r => clean((r||[])[k])).filter(Boolean);
          if(vals.length && vals.filter(isAlpha).length / vals.length > .6){ out.push(k); k++; } else break;
        } return out; };
      apeCols = extend(apeCols); nomCols = extend(nomCols);
    }
    if(!apeCols.length && !nomCols.length) fullCol = map.full != null ? map.full : null;

    // Matrícula: SIEMPRE de la columna indicada (por defecto AI), fila por fila
    const codeCol = opts.codeCol;
    if(codeCol != null) codeLabel = L(codeCol) + (hdr[codeCol] ? " (" + hdr[codeCol] + ")" : "");
    for(const row of data){
      if(!row) continue;
      pushStudent(row, codeCol, {
        anio: map.anio != null ? clean(row[map.anio]) : "",
        sec: map.sec != null ? clean(row[map.sec]) : "",
        grado: map.grado != null ? clean(row[map.grado]) : "",
        par: map.par != null ? clean(row[map.par]) : "",
        inst: map.inst != null ? clean(row[map.inst]) : ""
      });
    }
  } else {
    // 2) Sin encabezados: apellidos en E,F y nombres en G,H (o las columnas indicadas)
    apeCols = opts.manual ? (opts.apeCols || []) : [4,5];
    nomCols = opts.manual ? (opts.nomCols || []) : [6,7];
    for(const row of rows){
      if(!row) continue;
      const a = joinCols(row, apeCols), n = joinCols(row, nomCols);
      if(!(a || n) || (a && !isAlpha(a)) || (n && !isAlpha(n))) continue;
      if(classify(row[apeCols[0]]) || classify(row[nomCols[0]])) continue;
      pushStudent(row, opts.codeCol);
      if(opts.codeCol != null) codeLabel = L(opts.codeCol);
    }
    // 3) Último recurso: código tipo 2019-0001 y el texto más largo como nombre completo
    if(!opts.manual && !students.length){
      apeCols = []; nomCols = [];
      for(const row of rows){
        if(!row) continue;
        const cells = row.map(clean);
        const ci = cells.findIndex(v => CODE_RE.test(v));
        let best = "", bi = -1;
        cells.forEach((v,i) => { const s = stripDots(v); if(i !== ci && /^[A-Za-zÁÉÍÓÚÑÜáéíóúñü' ]+$/.test(s) && s.split(" ").length >= 2 && s.length > best.length){ best = s; bi = i; } });
        if(bi < 0 || (ci < 0 && students.length === 0 && /unidad|secci|grado|profesor|asignatura|a[ñn]o/i.test(best))) continue;
        if(BAD.test(best) || /unidad educativa|secci[oó]n|educaci[oó]n general/i.test(best)) continue;
        fullCol = bi;
        students.push({code: ci >= 0 ? cells[ci] : "", ape:"", nom:"", full: best.toUpperCase(), anio:"", sec:"", grado:"", par:"", inst:""});
        if(ci >= 0) codeLabel = "(códigos tipo 2019-0001)";
      }
    }
  }
  return {meta, students, headers, codeLabel, apeCols, nomCols, fullCol, hdr: hRow >= 0 ? (rows[hRow]||[]).map(clean) : []};
}

function buildGroups(parsed){
  const out = new Map(); let inst = "";
  for(const {meta, students} of parsed){
    if(meta.inst && !inst) inst = meta.inst;
    for(const s of students){
      if(s.inst && !inst) inst = s.inst;
      const g = {
        anio: fmtAnio(s.anio || meta.anio || ""),
        sec: fmtSec(s.sec || meta.sec || ""),
        grado: fmtGrado(s.grado || meta.grado || ""),
        par: fmtPar(s.par || meta.par || "")
      };
      const key = [g.anio,g.sec,g.grado,g.par].join("|");
      if(!out.has(key)) out.set(key, Object.assign(g, {key, students: []}));
      out.get(key).students.push({code: s.code, ape: s.ape, nom: s.nom, full: s.full});
    }
  }
  return {groups: [...out.values()], inst};
}

/* Orden del nombre: "ape" = apellidos primero, "nom" = nombres primero */
let nameOrder = "ape";
try{ if(localStorage.getItem("nomina.order") === "nom") nameOrder = "nom"; }catch(_){}
function displayName(s){
  if(s.full && !s.ape && !s.nom) return s.full;
  return (nameOrder === "ape" ? [s.ape, s.nom] : [s.nom, s.ape]).filter(Boolean).join(" ");
}

/* ---------- Render ---------- */
function shortLabel(g){
  const gr = g.grado.replace(/\s*grado de educaci[oó]n general b[aá]sica/i," EGB");
  return [gr, g.par.replace(/paralelo\s*/i,"")].filter(Boolean).join(" ") || "Curso";
}

function render(){
  const stage = $("stage");
  stage.querySelectorAll(".sheet").forEach(n => n.remove());
  const inst = clean($("fInst").value);
  const anioOv = clean($("fAnio").value);
  const trim = clean($("fTrim").value);
  const asig = clean($("fAsig").value);
  const guia = clean($("fGuia").value);
  const prof = clean($("fProf").value);
  let nCols = parseInt($("fCols").value, 10); if(isNaN(nCols)) nCols = 11; nCols = Math.max(0, Math.min(25, nCols));
  document.querySelectorAll("#presets button").forEach(b => b.setAttribute("aria-pressed", String(+b.dataset.n === nCols)));
  const sorted = $("fSort").checked;

  const colW = nCols ? Math.min(7.6, 100 / nCols) : 0;         // mm
  const nW = 6, leftW = 190 - nW - nCols * colW;                 // mm (ancho útil 190mm)
  const showCode = groups.some(g => g.students.some(s => s.code));

  const frag = document.createDocumentFragment(); let pages = 0;
  groups.forEach((g, gi) => {
    if(!passes(g)) return;
    let list = g.students.map(s => ({code: s.code, name: displayName(s)}));
    if(sorted) list.sort((a,b) => a.name.localeCompare(b.name, "es"));
    const chunks = []; for(let i = 0; i < Math.max(1, list.length); i += ROWS_PER_PAGE) chunks.push(list.slice(i, i + ROWS_PER_PAGE));
    chunks.forEach((chunk, ci) => {
      const start = ci * ROWS_PER_PAGE;
      const sec = document.createElement("section"); sec.className = "sheet";
      const cols = `<col style="width:${nW}mm"><col style="width:${leftW}mm">` + `<col style="width:${colW}mm">`.repeat(nCols);
      const ghs = `<th class="gh"></th>`.repeat(nCols);
      const rowsHtml = chunk.map((s, i) => `<tr><td class="n">${start + i + 1}</td><td class="nm"><div class="nmw">${showCode ? `<span class="cd">${esc(s.code)}</span>` : ""}<span class="name">${esc(s.name)}</span><span class="lead">${". ".repeat(60)}</span></div></td>${`<td class="g"></td>`.repeat(nCols)}</tr>`).join("");
      const grLine = [g.grado, g.par].filter(Boolean).join("  ");
      sec.innerHTML = `
        <div class="top"><img class="logo" src="${LOGO}" alt="Escudo de la Unidad Educativa Domingo Celi"><div class="inst">${esc(inst)}</div><div class="anio">Año lectivo : ${esc(anioOv || g.anio || "__________")}</div></div>
        <table>
          <colgroup>${cols}</colgroup>
          <thead><tr>
            <th class="hbox" colspan="2">
              <div class="hsec">${esc(g.sec || "SECCIÓN ________")}</div>
              <div class="hgr">${esc(grLine || "Grado ____________  Paralelo ___")}</div>
              <div class="fld"><span class="lbl">Trimestre :</span><span class="ln short">${esc(trim)}</span></div>
              <div class="fld"><span class="lbl">Asignatura :</span><span class="ln mid">${esc(asig)}</span></div>
              <div class="fld"><span class="lbl">Profesor Guía :</span><span class="ln long">${esc(guia)}</span></div>
            </th>${ghs}
          </tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
        ${ci === chunks.length - 1
          ? `<div class="sign"><span>Nombre Profesor: <span class="ln">${esc(prof)}</span></span><span>f.: <span class="ln"></span></span></div>`
          : `<div class="cont">Continúa en la página siguiente</div>`}`;
      frag.appendChild(sec); pages++;
    });
  });
  stage.appendChild(frag);
  $("empty").hidden = groups.length > 0;
  $("btnPrint").disabled = groups.length === 0;
  if(!groups.length){ $("pageCount").textContent = "Vista previa"; return; }
  const shown = groups.filter(passes);
  const total = shown.reduce((n,g) => n + g.students.length, 0);
  $("pageCount").textContent = `${pages} ${pages === 1 ? "hoja" : "hojas"} · ${total} ${total === 1 ? "estudiante" : "estudiantes"} · ${shown.length} ${shown.length === 1 ? "curso" : "cursos"}` + (shown.length < groups.length ? ` (de ${groups.length})` : "");
}

/* ---------- Filtros por curso y paralelo ---------- */
let fCurso = "all", fPar = "all";
const cursoLabel = g => g.grado.replace(/\s*grado de educaci[oó]n general b[aá]sica/i,"").trim() || "Sin curso";
const parLabel = g => g.par.replace(/paralelo\s*/i,"").trim() || "Sin paralelo";
const passes = g => (fCurso === "all" || g.grado === fCurso) && (fPar === "all" || g.par === fPar);

function chipsHtml(items, current){
  const total = items.reduce((n,i) => n + i.count, 0);
  return [{value:"all", label:"Todos", count: total}, ...items].map(i =>
    `<button type="button" data-v="${esc(i.value)}" aria-pressed="${i.value === current}">${esc(i.label)} <small>${i.count}</small></button>`).join("");
}
function refreshGroupUI(){
  $("filters").hidden = groups.length === 0;
  // Cursos (en el orden del Excel)
  const cursos = new Map();
  groups.forEach(g => { const c = cursos.get(g.grado) || {value: g.grado, label: cursoLabel(g), count: 0}; c.count += g.students.length; cursos.set(g.grado, c); });
  if(fCurso !== "all" && !cursos.has(fCurso)) fCurso = "all";
  // Paralelos del curso elegido
  const pars = new Map();
  groups.filter(g => fCurso === "all" || g.grado === fCurso).forEach(g => { const p = pars.get(g.par) || {value: g.par, label: parLabel(g), count: 0}; p.count += g.students.length; pars.set(g.par, p); });
  if(fPar !== "all" && !pars.has(fPar)) fPar = "all";
  const parList = [...pars.values()].sort((a,b) => a.label.localeCompare(b.label, "es"));
  $("fltCurso").innerHTML = chipsHtml([...cursos.values()], fCurso);
  $("fltPar").innerHTML = chipsHtml(parList, fPar);
}
$("fltCurso").addEventListener("click", e => { const b = e.target.closest("button[data-v]"); if(!b) return; fCurso = b.dataset.v; refreshGroupUI(); render(); });
$("fltPar").addEventListener("click", e => { const b = e.target.closest("button[data-v]"); if(!b) return; fPar = b.dataset.v; refreshGroupUI(); render(); });

function setStatus(msg, warn){ const s = $("status"); s.textContent = msg; s.classList.toggle("warn", !!warn); }

/* ---------- Archivo ---------- */
let book = null;   // {name, sheets:[{name, rows}]}

const L = i => XLSX.utils.encode_col(i);
function lettersToCols(v){
  return (clean(v).toUpperCase().match(/[A-Z]{1,3}/g) || []).map(x => XLSX.utils.decode_col(x));
}
// Apellidos y nombres SIEMPRE de las columnas indicadas (por defecto E:F y G:H)
function currentOpts(firstLoad){
  let apeCols = lettersToCols($("fApe").value), nomCols = lettersToCols($("fNom").value);
  if(!apeCols.length && !nomCols.length){ apeCols = [4,5]; nomCols = [6,7]; $("fApe").value = "E, F"; $("fNom").value = "G, H"; }
  let codeCol = lettersToCols($("fMat").value)[0];
  if(codeCol == null){ codeCol = XLSX.utils.decode_col("AI"); $("fMat").value = "AI"; }
  return {codeCol, apeCols, nomCols, manual: true};
}
function loadBook(firstLoad){
  const parsed = book.sheets.map(s => parseSheet(s.rows, s.name, currentOpts(firstLoad)));
  const res = buildGroups(parsed);
  if(!res.groups.length){ setStatus(`No encontré estudiantes en "${book.name}". Revisa que los apellidos estén en las columnas ${$("fApe").value} y los nombres en ${$("fNom").value}.`, true); return; }
  groups = res.groups;
  if(firstLoad){
    fCurso = "all"; fPar = "all";
    if(res.inst) $("fInst").value = res.inst.toUpperCase();
    $("fAnio").value = ""; $("fAnio").placeholder = groups[0].anio || "Escríbelo aquí";
  }
  refreshGroupUI(); render();

  const total = groups.reduce((n,g) => n + g.students.length, 0);
  const codeLabel = parsed.map(p => p.codeLabel).find(Boolean) || $("fMat").value;
  const conCodigo = groups.reduce((n,g) => n + g.students.filter(s => s.code).length, 0);
  const hasCodes = conCodigo === total;
  const missing = [];
  if(!groups.some(g => g.anio)) missing.push("año lectivo");
  if(!groups.some(g => g.sec)) missing.push("sección");
  if(!groups.some(g => g.grado)) missing.push("grado");
  if(!groups.some(g => g.par)) missing.push("paralelo");
  let msg = `Leí ${total} ${total === 1 ? "estudiante" : "estudiantes"} en ${groups.length} ${groups.length === 1 ? "curso" : "cursos"} desde "${book.name}".`;
  msg += hasCodes ? ` Matrícula tomada de la columna ${codeLabel}.` : ` Matrícula de la columna ${codeLabel}: ${total - conCodigo} ${total - conCodigo === 1 ? "estudiante no tiene" : "estudiantes no tienen"} dato en esa columna.`;
  const first = parsed.find(p => p.students.length) || parsed[0];
  const split = first.apeCols.length || first.nomCols.length;
  if(split) msg += ` Apellidos: columnas ${first.apeCols.map(L).join(", ") || "—"} · Nombres: columnas ${first.nomCols.map(L).join(", ") || "—"}.`;
  else if(first.fullCol != null) msg += ` El nombre completo viene en una sola columna (${L(first.fullCol)}); indica las columnas de apellidos y nombres para poder cambiar el orden.`;
  if(missing.length) msg += ` Falta en el Excel: ${missing.join(", ")}.`;
  setStatus(msg, !hasCodes || missing.length > 0 || !split);
}

function handleFile(file){
  if(!file) return;
  if(typeof XLSX === "undefined"){ setStatus("No se pudo cargar el lector de Excel. Revisa tu conexión a internet y vuelve a abrir la página.", true); return; }
  const reader = new FileReader();
  reader.onload = e => {
    try{
      const wb = XLSX.read(new Uint8Array(e.target.result), {type:"array", cellDates:false});
      book = {name: file.name, sheets: wb.SheetNames.map(n => {
        const ws = wb.Sheets[n];
        // La hoja se lee siempre desde A1, para que las letras (E, F, G, H, AI) coincidan con las del Excel
        if(ws["!ref"]){ const r = XLSX.utils.decode_range(ws["!ref"]); r.s.c = 0; r.s.r = 0; ws["!ref"] = XLSX.utils.encode_range(r); }
        return {name: n, rows: XLSX.utils.sheet_to_json(ws, {header:1, defval:"", raw:false, blankrows:true})};
      })};
      loadBook(true);
    }catch(err){
      setStatus(`No pude leer "${file.name}". Comprueba que sea un archivo de Excel válido (.xlsx, .xls o .csv).`, true);
    }
  };
  reader.readAsArrayBuffer(file);
}
$("fMat").addEventListener("change", () => { if(book) loadBook(false); });
["fApe","fNom"].forEach(id => $(id).addEventListener("change", () => { if(book) loadBook(false); }));

/* ---------- Orden: apellidos o nombres primero ---------- */
function setOrder(o){
  nameOrder = o;
  $("ordApe").setAttribute("aria-pressed", String(o === "ape"));
  $("ordNom").setAttribute("aria-pressed", String(o === "nom"));
  try{ localStorage.setItem("nomina.order", o); }catch(_){}
  render();
}
$("ordApe").addEventListener("click", () => setOrder("ape"));
$("ordNom").addEventListener("click", () => setOrder("nom"));

/* ---------- Configurador de columnas ---------- */
function setCols(n){
  n = Math.max(0, Math.min(25, n | 0));
  $("fCols").value = n;
  try{ localStorage.setItem("nomina.fCols", String(n)); }catch(_){}
  render();
}
$("colsMinus").addEventListener("click", () => setCols((parseInt($("fCols").value,10) || 0) - 1));
$("colsPlus").addEventListener("click", () => setCols((parseInt($("fCols").value,10) || 0) + 1));
$("presets").addEventListener("click", e => { const b = e.target.closest("button[data-n]"); if(b) setCols(+b.dataset.n); });

$("file").addEventListener("change", e => { handleFile(e.target.files[0]); e.target.value = ""; });
const drop = $("drop");
["dragenter","dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave","drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", e => handleFile(e.dataTransfer.files[0]));

/* ---------- Campos (se recuerdan en este navegador) ---------- */
const SAVED = ["fInst","fTrim","fAsig","fGuia","fProf","fCols"];
SAVED.forEach(id => {
  try{ const v = localStorage.getItem("nomina." + id); if(v != null) $(id).value = v; }catch(_){}
  $(id).addEventListener("input", () => { try{ localStorage.setItem("nomina." + id, $(id).value); }catch(_){} render(); });
});
if(!clean($("fInst").value)) $("fInst").value = DEFAULT_INST;
["fAnio"].forEach(id => $(id).addEventListener("input", render));
["fSort"].forEach(id => $(id).addEventListener("change", render));

$("btnPrint").addEventListener("click", () => window.print());

$("btnTpl").addEventListener("click", () => {
  if(typeof XLSX === "undefined"){ setStatus("No se pudo cargar el lector de Excel. Revisa tu conexión a internet.", true); return; }
  const aoa = [["N°","Matrícula","Año lectivo","Sección","Apellido paterno","Apellido materno","Primer nombre","Segundo nombre","Grado","Paralelo"]];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{wch:5},{wch:11},{wch:11},{wch:10},{wch:16},{wch:16},{wch:14},{wch:14},{wch:8},{wch:8}];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Estudiantes");
  XLSX.writeFile(wb, "plantilla_nomina.xlsx");
});

/* ---------- Inicio: sin datos hasta subir el Excel ---------- */
groups = [];
refreshGroupUI(); setOrder(nameOrder);
})();
</script>
</body>
</html>
