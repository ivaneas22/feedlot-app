const [scale, setScale] = useState(1);
const viewportRef = useRef(null);
const contentRef = useRef(null);

const fitToScreen = () => {
  if (!viewportRef.current || !contentRef.current) return;
  const vw = viewportRef.current.clientWidth;    // ancho visible
  const cw = contentRef.current.scrollWidth;     // ancho real del contenido
  if (!cw) return;
  const s = Math.max(0.5, Math.min(2, vw / cw)); // limita entre 50% y 200%
  setScale(Number(s.toFixed(2)));
};

import React, { useMemo, useState, useEffect } from "react";
import JSZip from "jszip";
import "./App.css";

/* ---------- Helpers de Excel (OpenXML via JSZip) ---------- */
function colLetterToIndex(col){let n=0;for(let i=0;i<col.length;i++) n=n*26+(col.charCodeAt(i)-64);return n-1;}
function cellRefToRC(ref){const m=ref?.match?.(/([A-Z]+)(\d+)/); if(!m) return null; return { r:parseInt(m[2],10)-1, c:colLetterToIndex(m[1])}; }
function parseRangeRef(ref){ if(!ref) return null; const parts=ref.split(":"); return { s:cellRefToRC(parts[0]), e:cellRefToRC(parts[1])}; }
function getAttr(el,name){return el?.getAttribute?.(name)??null;}

function readSharedStrings(xml){
  if(!xml) return [];
  const doc=new DOMParser().parseFromString(xml,"application/xml");
  const ss=[];
  doc.querySelectorAll("si").forEach(si=>{
    const tmp=[]; si.querySelectorAll("t").forEach(t=>tmp.push(t.textContent||""));
    ss.push(tmp.join(""));
  });
  return ss;
}
function readStyles(xml){
  if(!xml) return {numFmts:new Map(),cellXfs:[]};
  const doc=new DOMParser().parseFromString(xml,"application/xml");
  const numFmts=new Map(); doc.querySelectorAll("numFmts numFmt").forEach(nf=>{numFmts.set(parseInt(getAttr(nf,"numFmtId"),10),getAttr(nf,"formatCode"));});
  const cellXfs=[]; doc.querySelectorAll("cellXfs xf").forEach(xf=>{cellXfs.push({numFmtId:parseInt(getAttr(xf,"numFmtId")||"0",10)})});
  return {numFmts,cellXfs};
}
function isDateFmt(fmtId,fmtMap,fmtCode){const std=new Set([14,15,16,17,18,19,20,21,22]); if(std.has(fmtId)) return true; const code=fmtCode||fmtMap.get(fmtId)||""; return /[dyhmse]/i.test(code);}
function excelDateToJS(serial){const epoch=new Date(Date.UTC(1899,11,30)); const ms=Math.round(Number(serial||0)*86400000); return new Date(epoch.getTime()+ms);}
function formatCellValue(v,t,s,shared,styles){
  if(t==="s"){const idx=parseInt(v||"0",10); return shared[idx]??"";}
  if(t==="b") return v==="1"?"TRUE":"FALSE";
  if(v==null) return "";
  const style=styles?.cellXfs?.[parseInt(s||"0",10)];
  const fmtId=style?.numFmtId??0;
  const code=styles?.numFmts?.get(fmtId)||"";
  const num=Number(v);
  if(!Number.isNaN(num)&&isDateFmt(fmtId,styles?.numFmts,code)) return excelDateToJS(num).toLocaleString();
  if(!Number.isNaN(num)){ if(code.includes("%")) return (num*100).toLocaleString(undefined,{maximumFractionDigits:2})+"%"; return num.toLocaleString(); }
  return String(v);
}
function rawCellValue(v,t,shared){
  if(t==="s"){const idx=parseInt(v||"0",10); return shared[idx]??"";}
  if(t==="b") return v==="1";
  if(v==null||v==="") return "";
  const n=Number(v);
  return Number.isNaN(n)?v:n;
}

async function bootstrapZip(file){
  const zip=await JSZip.loadAsync(file);
  const [wbXml,ssXml,stylesXml,relsXml]=await Promise.all([
    zip.file("xl/workbook.xml").async("text"),
    zip.file("xl/sharedStrings.xml")?.async("text").catch(()=>""),
    zip.file("xl/styles.xml")?.async("text").catch(()=>""),
    zip.file("xl/_rels/workbook.xml.rels").async("text")
  ]);
  const shared=readSharedStrings(ssXml);
  const styles=readStyles(stylesXml);

  const wbDoc=new DOMParser().parseFromString(wbXml,"application/xml");
  const sheets=[]; wbDoc.querySelectorAll("sheets sheet").forEach(sh=>{sheets.push({name:getAttr(sh,"name"), rId:getAttr(sh,"r:id")});});

  const relsDoc=new DOMParser().parseFromString(relsXml,"application/xml");
  const rmap=new Map(); relsDoc.querySelectorAll("Relationship").forEach(r=>{rmap.set(getAttr(r,"Id"), getAttr(r,"Target"));});

  const sheetEntries=sheets.map(s=>({name:s.name, path:"xl/"+rmap.get(s.rId)}));
  return {zip,shared,styles,sheetEntries};
}
function readSheet(xml, shared, styles){
  const doc = new DOMParser().parseFromString(xml, "application/xml");

  // 1) Leer celdas como antes
  const cells = new Map();
  doc.querySelectorAll("sheetData row c").forEach(c => {
    const r = getAttr(c, "r"), t = getAttr(c, "t"), s = getAttr(c, "s");
    const v = c.querySelector("v")?.textContent ?? "";
    const rc = cellRefToRC(r); if (!rc) return;
    const key = `${rc.r}:${rc.c}`;
    cells.set(key, {
      raw:  rawCellValue(v, t, shared),
      text: formatCellValue(v, t, s, shared, styles)
    });
  });

  // 2) PROPAGAR VALORES EN CELDAS FUSIONADAS (mergeCells)
  //    Si una celda (e.g. MIXER) está fusionada A2:A10, Excel guarda el valor SOLO en A2.
  //    Repetimos ese valor en todas las celdas del rango que estén vacías.
  const merges = Array.from(doc.querySelectorAll("mergeCells mergeCell")) || [];
  for (const mc of merges) {
    const ref = getAttr(mc, "ref"); // ej: "A2:A10" o "C5:E5"
    const rng = parseRangeRef(ref);
    if (!rng || !rng.s || !rng.e) continue;

    const topKey = `${rng.s.r}:${rng.s.c}`;
    const top = cells.get(topKey);
    if (!top) continue; // si ni la top tiene valor, no hay nada que propagar

    // valor a propagar
    const topRaw  = top.raw;
    const topText = top.text;

    for (let R = rng.s.r; R <= rng.e.r; R++) {
      for (let C = rng.s.c; C <= rng.e.c; C++) {
        const key = `${R}:${C}`;
        if (key === topKey) continue; // la top ya tiene valor
        const curr = cells.get(key);
        const isEmpty = !curr || (
          (curr.raw === "" || curr.raw == null) &&
          (curr.text === "" || curr.text == null)
        );
        if (isEmpty) {
          cells.set(key, { raw: topRaw, text: topText });
        }
      }
    }
  }

  return cells;
}
// Rellena una matriz [filas][columnas] desde el Map de celdas
function rowsFromRange(range, cells) {
  if (!range) return [];
  const out = [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    const row = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const key = `${R}:${C}`;
      const cell = cells.get(key) || { raw: "", text: "" };
      row.push(cell);
    }
    out.push(row);
  }
  return out;
}
// Convierte la matriz de celdas en objetos {headers, displayRows, rawRows}
function buildRowsObjects(headersRow, matrix){
  const headers = (headersRow || []).map((h, i) => (h?.text || h?.raw || `Col ${i+1}`));

  const displayRows = [];
  const rawRows = [];

  for (let i = 1; i < (matrix?.length || 0); i++) {
    const r = matrix[i] || [];
    const disp = {};
    const raw = {};

    for (let j = 0; j < headers.length; j++) {
      const key = headers[j];
      disp[key] = r[j]?.text ?? "";
      raw[key]  = r[j]?.raw  ?? "";
    }

    displayRows.push(disp);
    rawRows.push(raw);
  }

  return { headers, displayRows, rawRows };
}

// Busca el índice de columna cuyo encabezado (en la fila de headers) coincida con /MIXER/i
function findHeaderColIndex(cells, headerRowIndex, regex = /(^|\s)mixer(\s|$)/i, scanFrom = 0, scanTo = 200) {
  for (let C = scanFrom; C <= scanTo; C++) {
    const key = `${headerRowIndex}:${C}`;
    const cell = cells.get(key);
    const txt = String(cell?.text ?? cell?.raw ?? "").trim();
    if (txt && regex.test(txt)) return C;
  }
  return -1;
}

/* ---------- EXTRACTOR con fallback (Tables o por nombre de hoja) ---------- */
async function extractBothTables(file){
  const {zip,shared,styles,sheetEntries}=await bootstrapZip(file);

  const xmlToDoc = (xml) => new DOMParser().parseFromString(xml,"application/xml");

  let formulaRes=null, comidaRes=null;

  // Camino 1: usar Excel Tables si existen
  for (const se of (sheetEntries||[])){
    const relsPath=se.path.replace("worksheets/","worksheets/_rels/")+".rels";
    const relsFile=zip.file(relsPath);
    const sheetXml=await zip.file(se.path).async("text");
    const cells=readSheet(sheetXml,shared,styles);

    if (relsFile){
      const relsXml=await relsFile.async("text");
      const relsDoc=xmlToDoc(relsXml);
      const tRels=Array.from(relsDoc.querySelectorAll('Relationship[Type$="/table"]'));

      for(const tr of tRels){
        const target=tr.getAttribute("Target");
        const norm = target.startsWith("/")
          ? target.slice(1)
          : (target.startsWith("../") ? "xl/"+target.replace("../","") : "xl/worksheets/"+target);
        const tfile=zip.file(norm) || zip.file("xl/"+target.replace("../",""));
        if(!tfile) continue;

        const tXml=await tfile.async("text");
        const tDoc=xmlToDoc(tXml);
        const tEl=tDoc.querySelector("table"); if(!tEl) continue;

        const tName=(tEl.getAttribute("name")||tEl.getAttribute("displayName")||"").toLowerCase();
        const ref=tEl.getAttribute("ref");
        const range=parseRangeRef(ref);
        const matrix=rowsFromRange(range,cells);
        const {headers,displayRows,rawRows}=buildRowsObjects(matrix[0]||[],matrix);

        if(tName==="formula" && !formulaRes){
          formulaRes={sheetName:se.name, tableName:"Formula", headers, data:displayRows};
        }
        if(tName==="comida" && !comidaRes){
          // ---- alias más amplios para MIXER ----
          const synonyms={ Corral:["Corral","N° Corral","Corral N","Corral Numero","Corral Número"],
            Cab:["Cab","Cabezas","Cab.","Cabeza"],
            Kg:["Kg","Kg Totales","Kg Total","Kilos","Kg/Día","Kg Dia","Kg Totales (Día)"],
            Etapa:["Etapa","Dieta","Fase"],
            "KgTC/Dia":["KgTC/Dia","KgTC/Día","KgTC x Dia","KgTC x Día","KgTC_Dia","Kg/cab/día","Kg c/cab","Kg por cab","Kg/cab"],
            MIXER:["MIXER","Mixer","N° Mixer","Mixer N°","Mixer Nro","Nro Mixer","Carro","Carro Mixer","Carro Nº","Carro N°","Carro Nro"],
            Ajuste:["Ajuste","Ajuste %","Ajuste%","% Ajuste"],
            KgManiana:["Kg Mañana","Kg Maniana","Kg Manana","Kg AM"],
            KgTarde:["Kg Tarde","Kg PM"] };
          const pickKey=(row,key)=>{
            const keys=Object.keys(row||{});
            const cand=(synonyms[key]||[key]).map(c=>c.toString().trim().toLowerCase());
            return keys.find(k=>cand.includes((k||"").toString().trim().toLowerCase()));
          };
          const trimmedRaw=rawRows.length>0?rawRows.slice(0,rawRows.length-1):rawRows;
          const trimmedDisp=displayRows.length>0?displayRows.slice(0,displayRows.length-1):displayRows;

          const outRows=(trimmedRaw||[]).map((rawRow,idx)=>{
            const dispRow=trimmedDisp[idx]||{};
            const g={};
            const kCorral=pickKey(rawRow,"Corral"); g.Corral=kCorral?(dispRow[kCorral]??rawRow[kCorral]??""):"";
            const kCab=pickKey(rawRow,"Cab"); const cabRaw=kCab?Number(rawRow[kCab]):NaN; g.Cab=Number.isFinite(cabRaw)?Math.round(cabRaw):0;
            const kEtapa=pickKey(rawRow,"Etapa"); g.Etapa=kEtapa?(dispRow[kEtapa]??rawRow[kEtapa]??""):"";
            // MIXER: priorizar el valor crudo (número) y no el texto formateado "Mixer 2"
const kMixer = pickKey(rawRow, "MIXER");
let mx = "";
if (kMixer) {
  // 1) primero intento con el valor crudo (si escribiste 2, acá viene 2)
  mx = rawRow[kMixer] ?? "";
  // 2) si no vino nada crudo, uso el display como respaldo
  if (mx === "" || mx == null) mx = dispRow[kMixer] ?? "";
}

if (mx === "" || mx == null) {
  g.MIXER = "Sin Mixer";
} else if (typeof mx === "number" && Number.isFinite(mx)) {
  // si ya es número, usamos ese
  g.MIXER = String(Math.round(mx));
} else {
  // si viene texto tipo "Mixer 2", extraigo el número
  const mtxt = String(mx);
  const numOnly = mtxt.match(/^\d+(?:\.\d+)?$/);
  const anyNum  = mtxt.match(/(\d+)/);
  if (numOnly) {
    g.MIXER = String(Math.round(Number(numOnly[0])));
  } else if (anyNum) {
    g.MIXER = String(Math.round(Number(anyNum[1])));
  } else if (/sin\s*mixer/i.test(mtxt)) {
    g.MIXER = "Sin Mixer";
  } else {
    g.MIXER = "Sin Mixer";
  }
}

            const kAjuste=pickKey(rawRow,"Ajuste"); const ajRaw=kAjuste!=null?Number(String(rawRow[kAjuste]).toString().replace("%","")):NaN; g.Ajuste=Number.isFinite(ajRaw)?(ajRaw<=1?ajRaw*100:ajRaw):100;
            const kKg=pickKey(rawRow,"Kg"); let kgRaw=kKg?Number(rawRow[kKg]):NaN;
            const kM=pickKey(rawRow,"KgManiana"); const kT=pickKey(rawRow,"KgTarde"); const m=kM?Number(rawRow[kM]):NaN; const t=kT?Number(rawRow[kT]):NaN;
            if(!Number.isFinite(kgRaw) && (Number.isFinite(m)||Number.isFinite(t))) kgRaw=(m||0)+(t||0);
            const kKgTC=pickKey(rawRow,"KgTC/Dia"); let kgtcRaw=kKgTC?Number(rawRow[kKgTC]):NaN; if(!Number.isFinite(kgtcRaw)&&Number.isFinite(kgRaw)&&Number.isFinite(cabRaw)&&cabRaw!==0) kgtcRaw=kgRaw/cabRaw;
            g["KgTC/Dia"]=Number.isFinite(kgtcRaw)?kgtcRaw:0;
            g.Kg=Number.isFinite(kgRaw)?kgRaw:(Number.isFinite(kgtcRaw)&&Number.isFinite(cabRaw)?kgtcRaw*cabRaw:0);
            return g;
          });
          // === Leer exactamente la columna 'MIXER' de la Excel Table y completar las filas vacías ===
try {
  // 1) ubicar índice de columna cuyo header sea exactamente 'MIXER' en matrix[0]
  const headerRow = matrix[0] || [];
  let mixerColIndex = -1;
  for (let c = 0; c < headerRow.length; c++) {
    const htxt = String(headerRow[c]?.text ?? headerRow[c]?.raw ?? '').trim().toLowerCase();
    if (htxt === 'mixer') { mixerColIndex = c; break; }
  }

  // 2) si existe esa columna, completar filas que quedaron con 'Sin Mixer'
  if (mixerColIndex >= 0) {
    for (let i = 0; i < outRows.length; i++) {
      if (outRows[i]?.MIXER !== 'Sin Mixer') continue;
      // matrix[0] es header, por eso i+1
      const cell = (matrix[i+1] || [])[mixerColIndex];
      const val  = cell?.raw ?? cell?.text ?? '';
      const m    = normalizeMixer(val);
      if (m !== 'Sin Mixer') outRows[i].MIXER = m;
    }
  }
} catch (_e) {
  // no-op
}

          // === Forzar lectura exacta desde la columna 'MIXER' de la tabla (camino Tables) ===
try {
  // 1) Ubicar índice de columna cuyo header sea exactamente 'MIXER'
  const headerRow = matrix[0] || [];
  let mixerColIndex = -1;
  for (let c = 0; c < headerRow.length; c++) {
    const htxt = String(headerRow[c]?.text ?? headerRow[c]?.raw ?? '').trim().toLowerCase();
    if (htxt === 'mixer') { mixerColIndex = c; break; }
  }

  // 2) Si existe esa columna, completar filas que quedaron con 'Sin Mixer'
  if (mixerColIndex >= 0) {
    for (let i = 0; i < outRows.length; i++) {
      if (outRows[i].MIXER !== 'Sin Mixer') continue;
      // matrix[i+1] porque matrix[0] es el header
      const cell = (matrix[i+1] || [])[mixerColIndex];
      const val  = cell?.raw ?? cell?.text ?? '';
      const m    = normalizeMixer(val);
      if (m !== 'Sin Mixer') outRows[i].MIXER = m;
    }
  }
} catch (_e) {
  // no-op
}

          // === Parche: si MIXER quedó "Sin Mixer", leerlo directo de la hoja por encabezado "MIXER" ===
// headerRowAbs = fila absoluta (en la hoja) donde está el header de la tabla
// dataStartAbs  = primera fila de datos absoluta (header + 1)
try {
  const headerRowAbs = range?.s?.r ?? 0;
  const dataStartAbs  = headerRowAbs + 1;

  // localizar columna "MIXER" en la fila de encabezado real (en la hoja)
  const colMixerAbs = findHeaderColIndex(
    cells,
    headerRowAbs,
    /(^|\s)mixer(\s|$)/i,
    range?.s?.c ?? 0,
    range?.e?.c ?? ((range?.s?.c ?? 0) + 80)
  );

  if (colMixerAbs >= 0) {
    for (let i = 0; i < outRows.length; i++) {
      if (outRows[i]?.MIXER !== "Sin Mixer") continue;

      // fila absoluta correspondiente a outRows[i]
      const rAbs = dataStartAbs + i;
      const kAbs = `${rAbs}:${colMixerAbs}`;

      // leer valor crudo o de texto directamente del XML de la hoja
      const cell = cells.get(kAbs);
      const val  = cell?.raw ?? cell?.text ?? "";

      const m    = normalizeMixer(val);
      if (m !== "Sin Mixer") outRows[i].MIXER = m;
    }
  }
} catch(_e) {
  // no-op: si algo falla, seguimos con lo que ya había
}

          // Fallback: si MIXER quedó "Sin Mixer", escanear TODA la fila para detectar 1–20,
// evitando columnas típicas numéricas (cab, kg, ajuste, etc.). Si nada, heredar último visto.
{
  const evitar = /(cab|cabeza|kg|kgtc|total|ajuste|ms|mv|%|dia|día|am|pm)/i;
  let ultimo = null;

  for (let idx = 0; idx < outRows.length; idx++) {
    const row = outRows[idx];
    if (row.MIXER !== "Sin Mixer") { ultimo = row.MIXER; continue; }

    const disp = trimmedDisp[idx] || {};
    const raw  = trimmedRaw[idx]  || {};
    let hallado = null;

    // 1) Probar columnas con nombres que contengan “mix/mixer/carro…”
    const candidatos1 = [
      ...Object.entries(disp),
      ...Object.entries(raw),
    ].filter(([k, v]) => {
      const kk = String(k||"");
      const vv = String(v||"");
      return /mix|mixer|carro|nro|n°|nº|#/.test(kk.toLowerCase()) || /mix|mixer|carro/.test(vv.toLowerCase());
    });

    for (const [, val] of candidatos1) {
      const m = normalizeMixer(val);
      if (m !== "Sin Mixer") { hallado = m; break; }
    }

    // 2) Si no apareció, escanear TODAS las celdas de la fila, excluyendo columnas obvias
    if (!hallado) {
      const candidatos2 = [
        ...Object.entries(disp),
        ...Object.entries(raw),
      ].filter(([k]) => !evitar.test(String(k||"").toLowerCase()));

      for (const [, val] of candidatos2) {
        const m = normalizeMixer(val);
        // Aceptar 1..20 como mixer válido
        if (m !== "Sin Mixer" && Number(m) >= 1 && Number(m) <= 20) { hallado = m; break; }
      }
    }

    if (hallado) { row.MIXER = hallado; ultimo = hallado; }
    else if (ultimo) { row.MIXER = ultimo; }
  }
}

          
          // Llenado hacia abajo (hereda el último MIXER visto, global)
{
  let last = null;
  for (const row of outRows) {
    if (row.MIXER === "Sin Mixer" && last) row.MIXER = last;
    if (row.MIXER !== "Sin Mixer") last = row.MIXER;
  }
}

          // Paso 6: refuerzo de MIXER en Tables (probar otras columnas + heredar último visto)
{
  const _lastMixer = { val: null };
  for (let idx = 0; idx < outRows.length; idx++) {
    const row = outRows[idx];
    if (row.MIXER !== "Sin Mixer") { _lastMixer.val = row.MIXER; continue; }

    // 6.a) sondear cualquier columna que parezca "mixer" o "carro" (display y raw)
    const disp = trimmedDisp[idx] || {};
    const raw  = trimmedRaw[idx]  || {};
    const probeEntries = [
      ...Object.entries(disp),
      ...Object.entries(raw),
    ].filter(([k,v]) => {
      const kk = String(k||"").toLowerCase();
      const vv = String(v||"");
      return /mix|mixer|carro|nro|n°|nº|#/.test(kk) || /mix|mixer|carro/.test(vv);
    });

    for (const [, val] of probeEntries) {
      const m = normalizeMixer(val);
      if (m !== "Sin Mixer") { row.MIXER = m; _lastMixer.val = m; break; }
    }

    // 6.b) si sigue vacío, heredar el último visto globalmente
    if (row.MIXER === "Sin Mixer" && _lastMixer.val) {
      row.MIXER = _lastMixer.val;
    }
  }
}

          // Paso 4: completar MIXER vacío heredando el último visto por corral
const _lastMixerByCorral = new Map();
for (const row of outRows) {
  const key = String(row.Corral||'').trim();
  if (!key) continue;
  if (row.MIXER === 'Sin Mixer' && _lastMixerByCorral.has(key)) {
    row.MIXER = _lastMixerByCorral.get(key);
  }
  if (row.MIXER !== 'Sin Mixer') {
    _lastMixerByCorral.set(key, row.MIXER);
  }
}


          comidaRes={ sheetName:se.name, tableName:"Comida", headers:["Corral","Cab","Kg","Etapa","KgTC/Dia","MIXER","Ajuste"], data:outRows };
        }
      }
    }
    if (formulaRes && comidaRes) break;
  }

  // Camino 2 (fallback): sin Tables -> detectar por nombre de hoja y leer todo el rango
  if (!formulaRes || !comidaRes){
    for (const se of (sheetEntries||[])){
      const sheetXml=await zip.file(se.path).async("text");
      const cells=readSheet(sheetXml,shared,styles);

      let range = (()=>{ // por <dimension> o por máximo r/c
        const doc=new DOMParser().parseFromString(sheetXml,"application/xml");
        const dim = doc.querySelector("dimension")?.getAttribute("ref") || null;
        if (dim) return parseRangeRef(dim);
        let maxR=0,maxC=0;
        for (const key of cells.keys()){
          const [r,c]=key.split(":").map(Number);
          if (r>maxR) maxR=r; if (c>maxC) maxC=c;
        }
        return { s:{r:0,c:0}, e:{r:maxR,c:maxC} };
      })();

      const matrix=rowsFromRange(range,cells);

      // Heurística de cabecera
      let headerRowIdx=0;
      const looksLikeHeader = (row=[])=>{
        const texts=row.map(c=>String(c?.text||"").trim());
        const nonEmpty=texts.filter(t=>t!=="").length;
        const alpha=texts.filter(t=>/[A-Za-z]/.test(t)).length;
        return nonEmpty>=3 && alpha>=2;
      };
      for (let i=0;i<Math.min(matrix.length,30);i++){
        if (looksLikeHeader(matrix[i])) { headerRowIdx=i; break; }
      }

      const headersRow = matrix[headerRowIdx] || [];
      const slice = matrix.slice(headerRowIdx);
      const {headers,displayRows}=buildRowsObjects(headersRow, slice);

      const low = (se.name||"").toLowerCase();

      if (!formulaRes && /formula|fórmula/.test(low) && headers.length){
        formulaRes={ sheetName: se.name, tableName:"Formula*", headers, data: displayRows };
      }

      if (!comidaRes && /comida|racion|ración|corrales/.test(low) && headers.length){
        const names = headers.map(h=>String(h).toLowerCase());
        const findName = (arr)=> headers[names.findIndex(n=>arr.includes(n))] || null;
        const K = {
          corral: findName(["corral","n° corral","nro corral","corral n","corral número","corral numero"]),
          cab: findName(["cab","cabezas","cab.","cabeza"]),
          etapa: findName(["etapa","dieta","fase"]),
          mixer: findName(["mixer","mezcladora","carro","carro mixer","carro nº","carro n°","carro nro","n° mixer","mixer n°","mixer nro","nro mixer"]),
          ajuste: findName(["ajuste","% ajuste","ajuste %","ajuste%","porcentaje ajuste"]),
          kgtc: findName(["kgtc/dia","kgtc/día","kg c/cab","kg por cab","kg/cab"]),
          kgtot: findName(["kg","kg totales","kg total","kilos","kg/día","kg dia","kg totales (día)"]),
          kgam: findName(["kg mañana","kg maniana","kg manana","kg am"]),
          kgpt: findName(["kg tarde","kg pm"]),
        };
        const out = displayRows.map((r)=> {
          const Cab = Number(r[K.cab] ?? 0);
          let KgTot = Number(r[K.kgtot] ?? NaN);
          const KgAM = Number(r[K.kgam] ?? NaN);
          const KgPT = Number(r[K.kgpt] ?? NaN);
          if (Number.isNaN(KgTot) && (!Number.isNaN(KgAM)||!Number.isNaN(KgPT))) {
            KgTot=(Number.isFinite(KgAM)?KgAM:0)+(Number.isFinite(KgPT)?KgPT:0);
          }
          let KgTC = Number(r[K.kgtc] ?? NaN);
          if (Number.isNaN(KgTC) && Number.isFinite(KgTot) && Number.isFinite(Cab) && Cab>0) KgTC=KgTot/Cab;
          if (Number.isNaN(KgTot) && Number.isFinite(KgTC) && Number.isFinite(Cab)) KgTot=KgTC*Cab;
          const AjusteRaw = Number(String(r[K.ajuste]??"").replace("%",""));
          const Ajuste = Number.isFinite(AjusteRaw) ? (AjusteRaw<=1?AjusteRaw*100:AjusteRaw) : 100;
          const MixerV = r[K.mixer];
          const MIXER = normalizeMixer(MixerV);
          return {
            Corral: String(r[K.corral]??""),
            Cab: Number.isFinite(Cab)?Math.round(Cab):0,
            Etapa: String(r[K.etapa]??""),
            "KgTC/Dia": Number.isFinite(KgTC)?KgTC:0,
            Kg: Number.isFinite(KgTot)?KgTot:0,
            MIXER: normalizerMixer(r[K.mixer]),
            Ajuste
          };
        });
        // Llenado hacia abajo (hereda el último MIXER visto, global)
{
  let last = null;
  for (const row of out) {
    if (row.MIXER === "Sin Mixer" && last) row.MIXER = last;
    if (row.MIXER !== "Sin Mixer") last = row.MIXER;
  }
}

        // Paso 5: completar MIXER vacío heredando el último visto por corral (fallback)
const _lastMixerByCorral_fb = new Map();
for (const row of out) {
  const key = String(row.Corral||'').trim();
  if (!key) continue;
  if (row.MIXER === 'Sin Mixer' && _lastMixerByCorral_fb.has(key)) {
    row.MIXER = _lastMixerByCorral_fb.get(key);
  }
  if (row.MIXER !== 'Sin Mixer') {
    _lastMixerByCorral_fb.set(key, row.MIXER);
  }
}

        comidaRes={ sheetName: se.name, tableName:"Comida*", headers:["Corral","Cab","Kg","Etapa","KgTC/Dia","MIXER","Ajuste"], data: out };
      }

      if (formulaRes && comidaRes) break;
    }
  }

  return { formulaRes, comidaRes };
}
/* ---------- UI: estilos por etapa ---------- */
const STYLES = {
  Inicio: {
    bg: 'bg-yellow-400',      // fondo intenso
    border: 'border-yellow-600', // borde más oscuro
    text: 'text-black'        // texto negro para buen contraste
  },
  Recria: {
    bg: 'bg-blue-500',
    border: 'border-blue-700',
    text: 'text-white'        // texto blanco para fondo oscuro
  },
  Terminacion: {
    bg: 'bg-red-500',
    border: 'border-red-700',
    text: 'text-white'
  },
  Default: {
    bg: 'bg-gray-100',
    border: 'border-gray-300',
    text: 'text-black'
  }
};


const stageNormUI = (etapaRaw) => {
  const s = String(etapaRaw || '').toLowerCase();
  if (/inicio/.test(s)) return 'Inicio';
  if (/recr/.test(s))  return 'Recria';
  if (/termin/.test(s)) return 'Terminacion';
  return 'Default';
};

/* --- Normalizador de MIXER robusto --- */
function normalizeMixer(v) {
  if (v == null) return "Sin Mixer";
  const raw = String(v).trim();
  if (raw === "" || /^0+$/.test(raw)) return "Sin Mixer";

  const s = raw
    .normalize("NFKD")
    .replace(/\s+/g, " ")
    .toLowerCase();

  // 1) Números explícitos con o sin ordinal (1, 2°, 03, “#2”, “Nº 4”, “N°5”, “No. 6”, “Mixer-07”)
  const numPat = /(?:^|[\s#:;,\-_/()])(?:no\.?|nº|n°|#)?\s*(\d{1,2})(?:[º°])?(?=$|[\s;,\-_/()])/i;
  const mNum = s.match(numPat);
  if (mNum) {
    const n = Math.round(Number(mNum[1]));
    return n > 0 ? String(n) : "Sin Mixer";
  }

  // 2) Romanos tras palabras clave (mix/mixer/carro)
  const romanPat = /(mix|mixer|carro|carrito)?\s*([ivx]{1,4})(?=$|[\s;,\-_/()])/i;
  const mRom = s.match(romanPat);
  if (mRom) {
    const roman = (mRom[2] || "").toUpperCase();
    const map = { I: 1, V: 5, X: 10 };
    let val = 0, prev = 0;
    for (let i = roman.length - 1; i >= 0; i--) {
      const cur = map[roman[i]] || 0;
      val += cur < prev ? -cur : cur;
      prev = cur;
    }
    if (val > 0 && val <= 50) return String(val);
  }

  // 3) Palabras (“uno”, “dos”, etc.)
  const words = {
    "uno":1,"dos":2,"tres":3,"cuatro":4,"cinco":5,
    "seis":6,"siete":7,"ocho":8,"nueve":9,"diez":10
  };
  const mWord = s.match(/\b(uno|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez)\b/i);
  if (mWord) {
    const n = words[mWord[1].toLowerCase()];
    if (n > 0) return String(n);
  }

  // 4) “sin mixer”
  if (/sin\s*mixer/.test(s)) return "Sin Mixer";

  // 5) Último intento: cualquier dígito aislado
  const fallback = s.match(/(^|\D)(\d{1,2})(\D|$)/);
  if (fallback) {
    const n = Math.round(Number(fallback[2]));
    return n > 0 ? String(n) : "Sin Mixer";
  }

  return "Sin Mixer";
}


/* ---------- Helpers visuales ---------- */
const isPctHeader = (h) => /%|inicio|recria|recría|termin/i.test(String(h||""));
const fmtPercent = (v) => {
  const n = Number(String(v).replace(/[^0-9.,-]/g,'').replace(',','.'));
  if (!Number.isFinite(n)) return v ?? "";
  const base = n<=1 ? (n*100) : n;
  return base.toLocaleString(undefined,{maximumFractionDigits:2}) + "%";
};
const fmtNumber = (v) => {
  const s = String(v ?? '');
  // extraigo solo dígitos, punto y coma
  const stripped = s.replace(/[^0-9.,-]/g, '').replace(',', '.');

  // si NO quedaron dígitos, entonces era texto puro (p. ej. "Silo") -> devuelvo el original
  if (!/[0-9]/.test(stripped)) return v ?? '';

  const n = Number(stripped);
  return Number.isFinite(n)
    ? n.toLocaleString(undefined, { maximumFractionDigits: 3 })
    : (v ?? '');
};

const fmtCell = (header, value) => isPctHeader(header) ? fmtPercent(value) : fmtNumber(value);

/* ---------- MixerTable ---------- */
function MixerTable({ info={}, insumoOrder=[], colorClass='' }) {
  const list = useMemo(() => {
    const keys = insumoOrder?.length ? insumoOrder : Object.keys(info?.insumosKg || {});
    return keys.filter(ing => {
      const val = (info?.insumosKg || {})[ing];
      if (!val || val <= 0) return false;

      const low = String(ing || '').toLowerCase().trim();

      // Excluir lo que NO es insumo real
      if (
        low === 'codigo' || low === 'código' ||
        low === 'dieta' ||
        low.startsWith('% ms') || low.startsWith('% mv')
      ) {
        return false;
      }

      return true;
    });
  }, [info, insumoOrder]);

  let acumulado = 0;
  return (
    <table className={`w-full border border-slate-200 rounded-xl shadow ${colorClass}`}>
      <thead className="bg-slate-50">
        <tr>
          <th className="px-3 py-2 text-left font-semibold border-b border-slate-200">Insumo</th>
          <th className="px-3 py-2 text-left font-semibold border-b border-slate-200">Kg Insumo</th>
          <th className="px-3 py-2 text-left font-semibold border-b border-slate-200">Kg Monitor</th>
        </tr>
      </thead>
      <tbody>
        {list.map(ing => {
          const kg = Math.round((info?.insumosKg || {})[ing] || 0);
          acumulado += kg;
          return (
            <tr key={ing} className="odd:bg-white even:bg-slate-50/70">
              <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap font-semibold">{ing}</td>
              <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">{kg.toLocaleString()}</td>
              <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">{acumulado.toLocaleString()}</td>
            </tr>
          );
        })}
        {/* Total visible al final */}
        <tr className="font-semibold">
          <td className="px-3 py-2 border-t border-slate-200">Total</td>
          <td className="px-3 py-2 border-t border-slate-200">
            {Math.round(info?.totalMixerKg || 0).toLocaleString()}
          </td>
          <td className="px-3 py-2 border-t border-slate-200">
            {Math.round(info?.totalMixerKg || 0).toLocaleString()}
          </td>
        </tr>
      </tbody>
    </table>
  );
}

/* ---------- MixerView (Drag & Drop) ---------- */
function MixerView({
  mixersAgg = {},
  insumoOrder = [],
  computedComida = [],
  mixerOrder = {},
  setMixerOrder = () => {},
}) {
  const handleDragStart = (fromIdx) => (e) => { e.dataTransfer.setData("text/plain", String(fromIdx)); };
  const handleDragOver = (e) => { e.preventDefault(); };
  const handleDrop = (mx, ordered) => (e) => {
    e.preventDefault();
    const fromIdx = Number(e.dataTransfer.getData("text/plain"));
    if (Number.isNaN(fromIdx)) return;
    const toIdx = Number(e.currentTarget.dataset.index);
    if (Number.isNaN(toIdx)) return;

    const next = ordered.map((x) => x.corral);
    const [moved] = next.splice(fromIdx, 1);
    next.splice(toIdx, 0, moved);

    setMixerOrder((prev) => ({ ...prev, [mx]: next }));
  };

  return (
    <div className="space-y-6">
      {Object.entries(mixersAgg || {}).map(([mx, info]) => {
        // Corrales con kg por ronda para este mixer
        const base = (computedComida || [])
          .filter((r) => String(r.MIXER) === String(mx))
          .map((r) => ({
            corral: String(r.Corral || ""),
            kg: Math.round(Number(r.KgRonda || 0)),
            etapa: stageNormUI(r.Etapa),
          }))
          .filter((x) => x.kg > 0);

        // Orden actual + faltantes al final
        const currentOrder = mixerOrder[mx] || [];
        const nameToItem = new Map(base.map((x) => [x.corral, x]));
        let ordered = currentOrder.map((n) => nameToItem.get(n)).filter(Boolean);
        const missing = base.filter((x) => !currentOrder.includes(x.corral));
        if (missing.length) ordered = [...ordered, ...missing];

        // Color por etapa dominante
        const counts = ordered.reduce((acc, r) => {
          if (r?.etapa && r.etapa !== "Default") acc[r.etapa] = (acc[r.etapa] || 0) + 1;
          return acc;
        }, {});
        const dominant = Object.entries(counts).sort((a, b) => b[1] - a[1])[0]?.[0] || "Default";
        const color = STYLES[dominant] || STYLES.Default;

        // Descarga acumulada
        let restante = Math.round(info?.totalMixerKg || 0);

        return (
          <div key={mx} className={`p-5 rounded-2xl shadow border ${color.bg} ${color.border}`}>
            <h3 className="text-lg font-semibold mb-3">MIXER {mx}</h3>

            {/* Lado a lado con alturas iguales (también al imprimir) */}
            <div className="grid grid-cols-2 gap-4 items-stretch">


              {/* Panel CARGA (izquierda) */}
              <div className="h-full flex flex-col flex-1">
                <div className="text-sm font-semibold mb-2">Carga por insumo</div>
                <div className="h-full">
                  
                  <MixerTable
                    info={info}
                    insumoOrder={insumoOrder}
                    colorClass={`${color.bg} h-full`}
                  />
                </div>
              </div>

              {/* Panel DESCARGA (derecha) */}
              <div className="h-full flex flex-col flex-1">
                <div className="text-sm font-semibold mb-2">Recorrido (arrastrá para reordenar)</div>
                <div className="h-full">
                  <table className={`w-full h-full border border-slate-200 rounded-xl text-sm ${color.bg}`}>
                    <thead className="bg-slate-50">
                      <tr>
                        <th className="px-3 py-2 text-left font-semibold border-b border-slate-200">Corral</th>
                        <th className="px-3 py-2 text-left font-semibold border-b border-slate-200">Kg por corral</th>
                        <th className="px-3 py-2 text-left font-semibold border-b border-slate-200">Descarga</th>
                      </tr>
                    </thead>
                    <tbody>
                      {ordered.map((c, idx) => {
                        restante = Math.max(0, restante - c.kg);
                        return (
                          <tr
                            key={c.corral}
                            data-index={idx}
                            draggable
                            onDragStart={handleDragStart(idx)}
                            onDragOver={handleDragOver}
                            onDrop={handleDrop(mx, ordered)}
                            className="cursor-move odd:bg-white even:bg-slate-50/70"
                            title="Arrastrá para cambiar el orden"
                          >
                            <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap font-semibold">{c.corral}</td>
                            <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">{c.kg.toLocaleString()}</td>
                            <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">{restante.toLocaleString()}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>
          </div>
        );
      })}
    </div>
  );
}

// Normaliza cualquier variante a: Inicio | Recria | Terminacion
function canonEtapa(v){
  const s = String(v||'')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,''); // saca tildes
  if (/inicio/.test(s))  return 'Inicio';
  if (/recr/.test(s))    return 'Recria';
  if (/termin/.test(s))  return 'Terminacion';
  return v || '';
}

/* ---------- App principal ---------- */
export default function App(){
  const [tab,setTab]=useState('formula');
  const [stateFormula,setStateFormula]=useState({loading:false,error:'',meta:null,rows:[]});
  const [stateComida,setStateComida]=useState({loading:false,error:'',meta:null,rows:[]});
  const [comidaEdit,setComidaEdit]=useState([]);
  const [mixerOrder,setMixerOrder]=useState({});

  const dietsFromFormula=useMemo(()=>{
    const defaults=["Inicio","Recria","Terminacion"];
    if(!(stateFormula?.rows||[]).length) return defaults;
    const keys=Object.keys(stateFormula.rows[0]||{});
    const guessCols=["Dieta","Etapa","Formula","Nombre","Tipo"];
    const col=keys.find(k=>guessCols.map(g=>g.toLowerCase()).includes(k.toLowerCase()));
    if(!col) return defaults;
    const vals=Array.from(new Set(stateFormula.rows.map(r=>String(r[col]||'').trim()).filter(Boolean)));
    const hits=vals.filter(v=>/inicio|recria|termin/i.test(v));
    return hits.length?hits:defaults;
  },[stateFormula]);

  const onFileBoth=async(e)=>{
    const file=e.target.files?.[0]; if(!file) return;
    setStateFormula({loading:true,error:'',meta:null,rows:[]});
    setStateComida({loading:true,error:'',meta:null,rows:[]});
    try{
      const {formulaRes,comidaRes}=await extractBothTables(file);

      if(formulaRes) setStateFormula({loading:false,error:'',meta:formulaRes,rows:formulaRes.data});
      else setStateFormula({loading:false,error:"No se encontró la tabla/hoja de 'Formula'",meta:null,rows:[]});

      if(comidaRes){
        setStateComida({loading:false,error:'',meta:comidaRes,rows:comidaRes.data});
        setComidaEdit((comidaRes.data||[]).map((r,i)=>({
          id:i,
          Corral:String(r.Corral??''),
          Cab:Number(r.Cab||0),
          Etapa:String(r.Etapa??''),
          MIXER: normalizeMixer(r.MIXER),
          Ajuste:Number(r.Ajuste||100),
          KgTC:Number(r['KgTC/Dia']||0)
        })));
      } else {
        setStateComida({loading:false,error:"No se encontró la tabla/hoja de 'Comida'",meta:null,rows:[]});
        setComidaEdit([]);
      }
    }catch(err){
      const msg=err?.message||String(err);
      setStateFormula({loading:false,error:msg,meta:null,rows:[]});
      setStateComida({loading:false,error:msg,meta:null,rows:[]});
      setComidaEdit([]);
    }
  };

  const updateRow=(id,patch)=>setComidaEdit(prev=>prev.map(r=>r.id===id?{...r,...patch}:r));

  const computedComida=useMemo(()=> (comidaEdit||[]).map(r=>{
    const cab=Number(r.Cab||0);
    const kgTC=Number(r.KgTC||0);
    const ajustePct=Number(r.Ajuste||100)/100;
    const totalDia=cab*kgTC*ajustePct;
    const kgRonda=totalDia/2;
    return {...r, TotalDia:totalDia, KgRonda:kgRonda};
  }),[comidaEdit]);

  // Inclusiones por etapa desde Fórmula (si existe)
  const inclusionMap=useMemo(()=>{
    const map={Inicio:{},Recria:{},Terminacion:{}};
    const rows=stateFormula?.rows||[]; if(!rows.length) return map;
    const headers=Object.keys(rows[0]||{});
    const stageCols=headers.filter(k=>/inicio|recria|recría|termin/i.test(k.toLowerCase()));
    let insumoKey=headers.find(k=>/insumo|ingred/i.test(k.toLowerCase()));
    if(!insumoKey){
      const nonNum=(k)=>{let c=0,t=0; rows.forEach(r=>{const v=r[k]; if(v==null||v==='') return; t++; const num=Number(String(v).replace(/[^0-9.,-]/g,'').replace(',','.')); if(Number.isNaN(num)||String(v).match(/[a-zA-Z]/)) c++;}); return t?(c/t):0;};
      const cand=headers.filter(h=>!stageCols.includes(h));
      insumoKey=cand.sort((a,b)=>nonNum(b)-nonNum(a))[0]||cand[0];
    }
    if(stageCols.length>=1 && (insumoKey in rows[0])){
      rows.forEach(r=>{
        const ing=String(r[insumoKey]||'').trim(); if(!ing) return;
        stageCols.forEach(sc=>{
          let num=Number(String(r[sc]).replace(/[^0-9.,-]/g,'').replace(',','.'));
          if(Number.isNaN(num)) return;
          if(num>1) num=num/100;
          const sk=/inicio/i.test(sc)?'Inicio':/recr/i.test(sc)?'Recria':'Terminacion';
          map[sk][ing]=num;
        });
      });
    }
    return map;
  },[stateFormula]);

  const insumoOrder=useMemo(()=>{
    const rows=stateFormula?.rows||[]; if(!rows.length) return [];
    const headers=Object.keys(rows[0]||{});
    const stageCols=headers.filter(k=>/inicio|recria|recría|termin/i.test(k.toLowerCase()));
    let insumoKey=headers.find(k=>/insumo|ingred/i.test(k.toLowerCase()));
    if(!insumoKey){
      const nonNum=(k)=>{let c=0,t=0; rows.forEach(r=>{const v=r[k]; if(v==null||v==='') return; t++; const num=Number(String(v).replace(/[^0-9.,-]/g,'').replace(',','.')); if(Number.isNaN(num)||String(v).match(/[a-zA-Z]/)) c++;}); return t?(c/t):0;};
      const cand=headers.filter(h=>!stageCols.includes(h));
      insumoKey=cand.sort((a,b)=>nonNum(b)-nonNum(a))[0]||cand[0];
    }
    return Array.from(new Set(rows.map(r=>String(r[insumoKey]||'').trim()).filter(Boolean)));
  },[stateFormula]);

  const mixersAgg=useMemo(()=>{
    const map={};
    (computedComida||[]).forEach(r=>{
      const mixer=(r.MIXER&&r.MIXER!=='Sin Mixer')?String(r.MIXER):null;
      if(!mixer) return;
      const etapa=stageNormUI(r.Etapa); if(etapa==='Default') return;
      const kgRonda=Number(r.KgRonda||0); if(!kgRonda) return;
      if(!map[mixer]) map[mixer]={ totalMixerKg:0, insumosKg:{} };
      map[mixer].totalMixerKg+=kgRonda;
      const inc=inclusionMap[etapa]||{};
      Object.entries(inc).forEach(([insumo,frac])=>{
        const low=String(insumo||'').toLowerCase();
        if(low==='total'||low.startsWith('% ms')||low.startsWith('% mv')) return;
        const add=kgRonda*Number(frac||0);
        map[mixer].insumosKg[insumo]=(map[mixer].insumosKg[insumo]||0)+add;
      });
    });
    return map;
  },[computedComida,inclusionMap]);

  useEffect(()=>{
    const next={...mixerOrder};
    Object.keys(mixersAgg||{}).forEach(mx=>{
      if(!next[mx]){
        const base=(computedComida||[]).filter(r=>String(r.MIXER)===String(mx)).map(r=>String(r.Corral||'')).filter(Boolean);
        if(base.length) next[mx]=base;
      }
    });
    if (JSON.stringify(next)!==JSON.stringify(mixerOrder)) setMixerOrder(next);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  },[mixersAgg,computedComida]);

  return (
    <div className="min-h-screen bg-slate-50 p-6">
      {/* Header tipo dashboard */}
<header className="w-full bg-white/90 backdrop-blur-sm border-b border-emerald-100 shadow-sm sticky top-0 z-10">
  <div className="max-w-7xl mx-auto flex justify-between items-center py-3 px-6">
    <h1 className="text-2xl font-bold bg-gradient-to-r from-emerald-600 to-emerald-400 bg-clip-text text-transparent">
      Feedlot Manager
    </h1>
    <span className="text-sm text-emerald-700 font-medium">
      Panel de Control
    </span>
  </div>
</header>
      <div className="max-w-7xl mx-auto space-y-5">

        {/* Título */}
        <h1 className="text-3xl md:text-4xl font-extrabold tracking-tight bg-gradient-to-r from-slate-900 to-slate-600 bg-clip-text text-transparent">
          Visor/Editor — Fórmula, Comida y Mixer
        </h1>

        {/* Uploader */}
        <label className="inline-flex items-center gap-3 px-4 py-2 rounded-2xl shadow bg-white border border-slate-200 cursor-pointer">
          <input type="file" accept=".xlsx,.xlsm" className="hidden" onChange={onFileBoth} />
          <span className="inline-flex items-center justify-center w-6 h-6 rounded-full bg-indigo-100 text-indigo-700 font-extrabold">↑</span>
          <span className="font-semibold">Seleccionar archivo XLSX</span>
        </label>

        {/* Tabs */}
        <div className="flex flex-wrap gap-2">
          <button onClick={()=>setTab('formula')} aria-current={tab==='formula'} className={`px-3 py-2 rounded-xl border shadow text-sm transition
              ${tab==='formula'
                ? 'bg-gradient-to-b from-indigo-50 to-indigo-100 border-indigo-200 text-indigo-900'
                : 'bg-slate-100 border-slate-200 text-slate-800 hover:-translate-y-0.5'}`}>
            Tabla Fórmula
          </button>
          <button onClick={()=>setTab('comida')} aria-current={tab==='comida'} className={`px-3 py-2 rounded-xl border shadow text-sm transition
              ${tab==='comida'
                ? 'bg-gradient-to-b from-indigo-50 to-indigo-100 border-indigo-200 text-indigo-900'
                : 'bg-slate-100 border-slate-200 text-slate-800 hover:-translate-y-0.5'}`}>
            Tabla Comida
          </button>
          <button onClick={()=>setTab('mixer')} aria-current={tab==='mixer'} className={`px-3 py-2 rounded-xl border shadow text-sm transition
              ${tab==='mixer'
                ? 'bg-gradient-to-b from-indigo-50 to-indigo-100 border-indigo-200 text-indigo-900'
                : 'bg-slate-100 border-slate-200 text-slate-800 hover:-translate-y-0.5'}`}>
            Mixer
          </button>
        </div>

        {/* FORMULA */}
        {tab==='formula' && (
          <div className="space-y-4">
            {stateFormula.loading && (<div className="p-4 rounded-2xl bg-white border border-slate-200 shadow text-slate-700">Procesando archivo…</div>)}
            {stateFormula.error && (<div className="p-4 rounded-2xl bg-red-50 border border-red-200 shadow text-red-700">{stateFormula.error}</div>)}
            {stateFormula.meta && (
              <div className="p-4 rounded-2xl bg-white border border-slate-200 shadow text-sm">
                <div><strong>Hoja:</strong> {stateFormula.meta.sheetName}</div>
                <div><strong>Tabla:</strong> {stateFormula.meta.tableName}</div>
                <div><strong>Columnas:</strong> {stateFormula.meta.headers.filter(Boolean).join(', ')}</div>
              </div>
            )}
            {stateFormula.rows?.length>0 ? (
              <div className="overflow-auto rounded-2xl border border-slate-200 bg-white shadow">
                <table className="min-w-full text-sm">
                  <thead className="sticky top-0 bg-gradient-to-b from-slate-900 to-slate-700 text-white">
                    <tr>
                      {stateFormula.meta.headers.map((h,i)=>(
                        <th key={i} className="px-3 py-2 text-left font-semibold border-b border-slate-700">{h||`Col ${i+1}`}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                  {stateFormula.rows
  .filter(r => String(r.Etapa || '').toLowerCase() !== 'dieta')
  .map((row,idx)=>(
  <tr key={idx} className="odd:bg-white even:bg-slate-50">

                        {stateFormula.meta.headers.map((h,i)=>{
                          const v = row[h||`Col ${i+1}`];
                          let pretty = fmtCell(h, v);
if (/codigo/i.test(h)) {
  const n = Number(String(v).replace(/[^0-9.-]/g, ''));
  if (Number.isFinite(n)) pretty = Math.round(n).toString();
}

                          return (
                            <td key={i} className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">{pretty}</td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (<div className="text-slate-600">Subí un archivo para ver la tabla de Fórmula.</div>)}
          </div>
        )}

        {/* COMIDA */}
        {tab==='comida' && (
          <div className="space-y-4">
            {stateComida.loading && (<div className="p-4 rounded-2xl bg-white border border-slate-200 shadow text-slate-700">Procesando archivo…</div>)}
            {stateComida.error && (<div className="p-4 rounded-2xl bg-red-50 border border-red-200 shadow text-red-700">{stateComida.error}</div>)}
            {stateComida.rows?.length>0 ? (
              <div className="overflow-auto rounded-2xl border border-slate-200 bg-white shadow">
                <table className="min-w-full text-sm">
                  <thead className="sticky top-0 bg-gradient-to-b from-slate-900 to-slate-700 text-white">
                    <tr>
                      {['Corral','Cab','Etapa','KgTC/Dia','Ajuste (%)','MIXER','Total Día','Kg Ronda'].map((h,i)=>(
                        <th key={i} className="px-3 py-2 text-left font-semibold border-b border-slate-700">{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {(computedComida||[]).map((row)=>(
                      <tr key={row.id} className="odd:bg-white even:bg-slate-50">
                        <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">{row.Corral}</td>
                        <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">{row.Cab}</td>
                        <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">
  <select
    className="border border-slate-300 rounded-lg px-2 py-1 w-32"
    value={row.Etapa || ''}
    onChange={(e)=>updateRow(row.id,{ Etapa: e.target.value })}
  >
    <option value="">—</option>
    {dietsFromFormula.map(opt => (
      <option key={opt} value={opt}>{opt}</option>
    ))}
  </select>
</td>
                        <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">
                          <input type="number" step="0.5" className="border border-slate-300 rounded-lg px-2 py-1 w-28" value={Number(row.KgTC||0)} onChange={(e)=>updateRow(row.id,{KgTC: Math.round(Number(e.target.value||0)*2)/2})}/>
                        </td>
                        <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">
                          <input type="number" step="0.5" className="border border-slate-300 rounded-lg px-2 py-1 w-24" value={Number(row.Ajuste||100)} onChange={(e)=>updateRow(row.id,{Ajuste: Number(e.target.value)||100})}/>
                        </td>
                        <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">
                          <select className="border border-slate-300 rounded-lg px-2 py-1" value={row.MIXER} onChange={(e)=>updateRow(row.id,{MIXER:e.target.value})}>
                            {["Sin Mixer","1","2","3","4","5"].map(m=>(<option key={m} value={m}>{m}</option>))}
                          </select>
                        </td>
                        <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">
                          {(row.TotalDia??0).toLocaleString(undefined,{minimumFractionDigits:2, maximumFractionDigits:2})}
                        </td>
                        <td className="px-3 py-2 border-b border-slate-200 whitespace-nowrap">
                          {(row.KgRonda??0).toLocaleString(undefined,{minimumFractionDigits:2, maximumFractionDigits:2})}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (<div className="text-slate-600">Subí el archivo con la tabla/hoja Comida.</div>)}
          </div>
        )}

        {/* MIXER */}
        {tab==='mixer' && (
          <div className="rounded-2xl border border-slate-200 bg-white shadow p-4">
            <MixerView
              mixersAgg={mixersAgg}
              insumoOrder={insumoOrder}
              computedComida={computedComida}
              mixerOrder={mixerOrder}
              setMixerOrder={setMixerOrder}
            />
          </div>
        )}
      </div>
    </div>
  );
}
