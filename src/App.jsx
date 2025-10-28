import React, { useCallback, useRef, useState } from "react";
import JSZip from "jszip";

/* ===== Utils XML ===== */
const unesc = s => (s||"").replaceAll("&lt;","<").replaceAll("&gt;",">")
  .replaceAll("&amp;","&").replaceAll("&quot;","\"").replaceAll("&apos;","'");
const readText = async (zip, p) => zip.file(p)?.async("string") ?? null;
const loadZip = async f => JSZip.loadAsync(await f.arrayBuffer());

/* ===== A1 helpers ===== */
const colToIndex = col => { let n=0; for (let i=0;i<col.length;i++) n=n*26+(col.charCodeAt(i)-64); return n-1; };
const rc = ref => { const m=ref.match(/([A-Z]+)(\d+)/); return { r:parseInt(m[2],10)-1, c:colToIndex(m[1])}; };
const parseRange = ref => { const m=ref.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/); return { tl:rc(m[1]+m[2]), br:rc(m[3]+m[4]) }; };
const addr = (r,c) => { let n=c+1,s=""; while(n>0){const k=(n-1)%26; s=String.fromCharCode(65+k)+s; n=Math.floor((n-1)/26);} return s+(r+1); };

/* ===== sharedStrings ===== */
async function parseSharedStrings(zip){
  const xml = await readText(zip, "xl/sharedStrings.xml"); if(!xml) return [];
  const sis = [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>m[1]);
  return sis.map(si => [...si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)]
    .map(mm=>unesc(mm[1])).join(""));
}

/* ===== styles (numFmt, font/fill por xf) ===== */
async function parseStyles(zip){
  const xml = await readText(zip, "xl/styles.xml");
  if(!xml) return { xfs:[], fonts:[], fills:[], numFmts:new Map() };

  const numFmts = new Map([...xml.matchAll(/<numFmt[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g)]
    .map(m=>[+m[1], unesc(m[2])]));

  const fonts = [...xml.matchAll(/<font>([\s\S]*?)<\/font>/g)].map(m=>{
    const s=m[1]; return { bold:/<b\b/.test(s),
      color:(s.match(/<color[^>]*rgb="([0-9A-Fa-f]{8})"/)||[])[1] };
  });
  const fills = [...xml.matchAll(/<fill>([\s\S]*?)<\/fill>/g)].map(m=>{
    return { fg:(m[1].match(/<fgColor[^>]*rgb="([0-9A-Fa-f]{8})"/)||[])[1] };
  });
  const xfs = [...xml.matchAll(/<xf[^>]*?(?:numFmtId="(\d+)")?[^>]*?(?:fontId="(\d+)")?[^>]*?(?:fillId="(\d+)")?[^>]*?>/g)]
    .map(m=>({ numFmtId:m[1]?+m[1]:null, fontId:m[2]?+m[2]:null, fillId:m[3]?+m[3]:null }));

  return { xfs, fonts, fills, numFmts };
}
const rgb = a => (!a||a.length!==8) ? undefined : "#"+a.slice(2);
const styleForXf = (xf, styles) => {
  if(!xf) return {};
  const st={};
  if (xf.fillId!=null) { const f=styles.fills[xf.fillId]; if(f?.fg) st.backgroundColor=rgb(f.fg); }
  if (xf.fontId!=null){ const ft=styles.fonts[xf.fontId]; if(ft?.bold) st.fontWeight="700"; if(ft?.color) st.color=rgb(ft.color); }
  return st;
};
function formatByNumFmt(value, xf, styles){
  if(value==null || value==="") return "";
  const n = Number(value); if (Number.isNaN(n)) return String(value);
  const id = xf?.numFmtId ?? null;

  // Estándar de Excel: 9 = 0%, 10 = 0.00%
  if (id===9 || id===10) {
    const d = id===10 ? 2 : 0;
    return (n*100).toLocaleString("es-AR", {minimumFractionDigits:d, maximumFractionDigits:d}) + "%";
  }
  // Personalizados con % en formatCode
  const fmt = styles.numFmts.get(id);
  if (fmt && fmt.includes("%")) {
    const m = fmt.match(/0\.(0+)/); const d = m? m[1].length : 0;
    return (n*100).toLocaleString("es-AR",{minimumFractionDigits:d, maximumFractionDigits:d})+"%";
  }
  // Default numérico
  return n.toLocaleString("es-AR", { maximumFractionDigits: 6 });
}

/* ===== localizar tabla por displayName y su hoja ===== */
async function findTable(zip, displayName){
  const tablePaths = Object.keys(zip.files).filter(p=>p.startsWith("xl/tables/table") && p.endsWith(".xml"));
  let tablePath=null, ref=null;
  for (const p of tablePaths){
    const xml = await readText(zip, p);
    const dn = (xml.match(/displayName="([^"]+)"/)||[])[1];
    if (dn && dn.toLowerCase()===displayName.toLowerCase()){
      ref = (xml.match(/ref="([^"]+)"/)||[])[1];
      tablePath = p; break;
    }
  }
  if(!tablePath || !ref) throw new Error(`No encontré la tabla '${displayName}'.`);

  // ¿qué hoja referencia esa tabla?
  const wsRels = Object.keys(zip.files).filter(p=>p.startsWith("xl/worksheets/_rels/") && p.endsWith(".rels"));
  let sheetXmlPath=null;
  for (const rel of wsRels){
    const x = await readText(zip, rel);
    if (x && x.includes(tablePath.replace("xl/",""))) {
      sheetXmlPath = rel.replace("/_rels","").replace(".rels","");
      break;
    }
  }
  if(!sheetXmlPath) throw new Error("No pude resolver la hoja de la tabla.");
  return { ref, sheetXmlPath };
}

/* ===== leer celdas de una hoja (valor, tipo t, estilo s) ===== */
async function parseSheet(zip, sheetXmlPath){
  const xml = await readText(zip, sheetXmlPath);
  if(!xml) throw new Error("No pude leer: "+sheetXmlPath);
  const cells = new Map(); // addr -> {t,v,s}
  const re = /<c r="([A-Z]+\d+)"[^>]*?(?:t="([a-zA-Z]+)")?[^>]*?(?:s="(\d+)")?[^>]*?>([\s\S]*?)<\/c>/g;
  for (const m of xml.matchAll(re)){
    const a=m[1], t=m[2]||null, s=m[3]?+m[3]:null;
    const inner = m[4]||"";
    const mv = inner.match(/<v>([\s\S]*?)<\/v>/);
    let v=null;
    if (mv) v = unesc(mv[1]);
    else {
      const mt = inner.match(/<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/is>/);
      if (mt) v = unesc(mt[1]);
    }
    cells.set(a, {t,v,s});
  }
  return cells;
}

/* ===== construir rango 2D + estilos ===== */
function buildRange(cells, shared, ref){
  const R = parseRange(ref);
  const rows=[], sidx=[];
  for(let r=R.tl.r; r<=R.br.r; r++){
    const row=[], srow=[];
    for(let c=R.tl.c; c<=R.br.c; c++){
      const a = addr(r,c);
      const cell = cells.get(a);
      if(!cell){ row.push(null); srow.push(null); continue; }
      let v = cell.v;
      if (cell.t==='s' && v!=null) v = shared[parseInt(v,10)] ?? null;
      if (cell.t==='b' && v!=null) v = v==='1';
      row.push(v); srow.push(cell.s ?? null);
    }
    rows.push(row); sidx.push(srow);
  }
  return { rows, sidx };
}
const rowsToObjects = (rows, hdrs) => {
  const headers = (hdrs && hdrs.length===rows[0].length) ? hdrs : rows[0].map(h=>String(h??""));
  const data = rows.slice(1).map(r=>{ const o={}; headers.forEach((h,i)=>o[h]=r[i]); return o; });
  return { headers, data };
};

/* ===== App ===== */
export default function App(){
  const [tab, setTab] = useState("Formula"); // "Formula" | "Corrales/Racion"
  const [formula, setFormula] = useState(null); // {headers, data, sidx}
  const [comida, setComida] = useState(null);   // {headers, data}
  const [styles, setStyles] = useState(null);
  const fileRef = useRef(null);

  const onFile = useCallback(async (e)=>{
    const file = e.target.files?.[0]; if(!file) return;
    try{
      const zip = await loadZip(file);
      const shared = await parseSharedStrings(zip);
      const sty = await parseStyles(zip);
      setStyles(sty);

      // Formula
      const { ref:refF, sheetXmlPath:sheetF } = await findTable(zip, "Formula");
      const cellsF = await parseSheet(zip, sheetF);
      const { rows:rowsF, sidx:sidxF } = buildRange(cellsF, shared, refF);
      const headersF = rowsF[0].map(x=> String(x??""));
      const { data:dataF } = rowsToObjects(rowsF, headersF);
      setFormula({ headers: headersF, data: dataF, sidx: sidxF });

      // comida
      const { ref:refC, sheetXmlPath:sheetC } = await findTable(zip, "comida");
      const cellsC = await parseSheet(zip, sheetC);
      const { rows:rowsC } = buildRange(cellsC, shared, refC);
      const { headers:headersCraw, data:dataCraw } = rowsToObjects(rowsC);
      const wanted = ["Corral","Cab","Etapa","KgTC/Dia","MIXER","Ajuste"];
      const headersC = wanted.filter(w => headersCraw.includes(w));
      const dataC = dataCraw.map(r => { const o={}; headersC.forEach(h=>o[h]=r[h]); return o; });
      setComida({ headers: headersC, data: dataC });

    }catch(err){
      console.error(err);
      alert("Error al leer el Excel: " + (err?.message || String(err)));
    }
  },[]);

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      <div className="max-w-6xl mx-auto">
        <h1 className="text-2xl font-bold mb-3">Modelo Feedlot — Importación única</h1>
        <p className="text-sm text-gray-700 mb-4">
          Respeta formato de número de Excel (porcentajes) y aplica colores/negrita básicos.
        </p>

        <div className="flex items-center gap-3 mb-4">
          <input ref={fileRef} type="file" accept=".xlsx" onChange={onFile} />
          <div className="ml-auto flex rounded-xl overflow-hidden border">
            <button onClick={()=>setTab("Formula")} className={`px-4 py-2 ${tab==="Formula"?"bg-gray-900 text-white":"bg-white"}`}>Formula</button>
            <button onClick={()=>setTab("Corrales/Racion")} className={`px-4 py-2 ${tab==="Corrales/Racion"?"bg-gray-900 text-white":"bg-white"}`}>Corrales/Racion</button>
          </div>
        </div>

        {tab==="Formula" && formula && styles && (
          <div className="bg-white rounded-2xl shadow p-4">
            <h2 className="font-semibold mb-3">Formula</h2>
            <div className="overflow-auto">
              <table className="min-w-full text-sm border-collapse border border-gray-300">
                <thead>
                  <tr>
                    {formula.headers.map((h,i)=>(
                      <th key={i} className="border p-2 bg-gray-50 text-left">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {formula.data.map((row,ri)=>(
                    <tr key={ri}>
                      {formula.headers.map((h,ci)=>{
                        const xf = styles.xfs?.[ formula.sidx?.[ri+1]?.[ci] ?? null ]; // +1 (saltear header)
                        const style = styleForXf(xf, styles);
                        const text  = formatByNumFmt(row[h], xf, styles);
                        return <td key={ci} className="border p-2" style={style}>{text}</td>;
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab==="Corrales/Racion" && comida && (
          <div className="bg-white rounded-2xl shadow p-4">
            <h2 className="font-semibold mb-3">Corrales/Racion</h2>
            <div className="overflow-auto">
              <table className="min-w-full text-sm border-collapse border border-gray-300">
                <thead>
                  <tr>{comida.headers.map((h,i)=>(
                    <th key={i} className="border p-2 bg-gray-50 text-left">{h}</th>
                  ))}</tr>
                </thead>
                <tbody>
                  {comida.data.map((row,ri)=>(
                    <tr key={ri}>
                      {comida.headers.map((h,ci)=>(
                        <td key={ci} className="border p-2">{row[h] ?? ""}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
