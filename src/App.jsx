import React, { useState, useEffect } from "react";
import JSZip from "jszip";
import "./App.css";

/* =================== MARCO VISUAL POR ETAPA =================== */
// etapa → clase CSS (solo si hay valor; si está vacío no colorea)
function getEtapaClass(etapa) {
  if (!etapa || String(etapa).trim() === "") return ""; // sin color si está vacío
  const e = String(etapa).toLowerCase();
  if (e.includes("term")) return "stage--terminacion";
  if (e.includes("recr")) return "stage--recria";
  return "stage--inicio"; // cualquier otro valor => inicio (ámbar)
}

// Encabezados: primera letra en mayúscula
function prettyHeader(h) {
  const s = String(h ?? "");
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}

// Caja visual que pinta borde/fondo según etapa (no toca lógica)
function MixerBox({ etapa, titulo, children }) {
  const clase = getEtapaClass(etapa);
  return (
    <div className={`stage-wrap ${clase}`}>
      <div className="stage-inner mixer-card">
        {titulo ? <div className="stage-badge">{titulo}</div> : null}
        {children}
      </div>
    </div>
  );
}

/* =================== UTILIDADES =================== */
const xmlToDoc = (xmlStr) =>
  new window.DOMParser().parseFromString(xmlStr, "application/xml");

const norm = (s) =>
  (s || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // sin acentos
    .replace(/\s+/g, "") // sin espacios
    .toLowerCase();

function colLetterToIndex(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
  return n - 1;
}
function cellRefToRC(ref) {
  const m = ref?.match?.(/([A-Z]+)(\d+)/);
  if (!m) return null;
  return { r: parseInt(m[2], 10) - 1, c: colLetterToIndex(m[1]) };
}
function parseRangeRef(ref) {
  if (!ref) return null;
  const m = ref.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  if (!m) return null;
  return {
    r0: parseInt(m[2], 10) - 1,
    c0: colLetterToIndex(m[1]),
    r1: parseInt(m[4], 10) - 1,
    c1: colLetterToIndex(m[3]),
  };
}

/* =================== ESTILOS (FORMATO NUMÉRICO) =================== */
function parseStyles(stylesXml) {
  if (!stylesXml) return { xfToFmt: {}, numFmtIdToCode: {} };
  const doc = xmlToDoc(stylesXml);
  const numFmtIdToCode = {};
  const numFmts = doc.getElementsByTagName("numFmt");
  for (let i = 0; i < numFmts.length; i++) {
    const nf = numFmts[i];
    const id = nf.getAttribute("numFmtId");
    const code = nf.getAttribute("formatCode");
    if (id && code) numFmtIdToCode[id] = code;
  }
  const builtIns = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    14: "m/d/yy",
    22: "m/d/yy h:mm",
  };
  const xfToFmt = {};
  const cellXfs = doc.getElementsByTagName("cellXfs")[0];
  if (cellXfs) {
    const xfs = cellXfs.getElementsByTagName("xf");
    for (let i = 0; i < xfs.length; i++) {
      const xf = xfs[i];
      const nfid = xf.getAttribute("numFmtId") || "0";
      const code = numFmtIdToCode[nfid] ?? builtIns[nfid] ?? "General";
      xfToFmt[i] = { numFmtId: nfid, code };
    }
  }
  return { xfToFmt, numFmtIdToCode: { ...builtIns, ...numFmtIdToCode } };
}

const isPct = (fmt) => /%/.test(fmt || "");
function formatNumber(value, fmt) {
  if (isPct(fmt)) {
    const d = /0\.00%/.test(fmt) ? 2 : /0%/.test(fmt) ? 0 : 2;
    return `${(Number(value) * 100).toFixed(d)}%`;
  }
  if (/0\.00/.test(fmt)) return Number(value).toFixed(2);
  if (/(^|[^.])0(?![0-9])/.test(fmt)) return String(Math.round(Number(value)));
  return String(Number(value));
}
function formatCell(rawV, t, sIdx, styles) {
  if (rawV == null || rawV === "") return { display: "", kind: "empty" };
  if (t === "s") return { display: rawV, kind: "text" };
  if (t === "b") return { display: rawV === "1" ? "TRUE" : "FALSE", kind: "bool" };
  const fmt = styles?.xfToFmt?.[sIdx]?.code ?? "General";
  if (t == null || t === "n" || t === "d")
    return {
      display: formatNumber(rawV, fmt),
      kind: isPct(fmt) ? "percent" : "number",
    };
  return { display: String(rawV), kind: "text" };
}

/* =================== PARSEAR HOJA =================== */
function parseSheetXml(sheetXml, sstArray, styles) {
  const doc = xmlToDoc(sheetXml);
  const dim = doc.getElementsByTagName("dimension")[0]?.getAttribute("ref");
  const cells = Array.from(doc.getElementsByTagName("c"));
  let range = parseRangeRef(dim);
  if (!range) {
    let maxR = 0, maxC = 0;
    for (const c of cells) {
      const rc = cellRefToRC(c.getAttribute("r"));
      if (!rc) continue;
      if (rc.r > maxR) maxR = rc.r;
      if (rc.c > maxC) maxC = rc.c;
    }
    range = { r0: 0, c0: 0, r1: maxR, c1: maxC };
  }
  const rows = Array.from(
    { length: range.r1 - range.r0 + 1 },
    () => Array(range.c1 - range.c0 + 1).fill({ display: "", kind: "empty" })
  );
  for (const c of cells) {
    const ref = c.getAttribute("r");
    const t = c.getAttribute("t");
    const sIdx = parseInt(c.getAttribute("s") ?? "-1", 10);
    const vNode = c.getElementsByTagName("v")[0];
    const fNode = c.getElementsByTagName("f")[0];
    let rawV = vNode?.textContent ?? "";
    let typ = t;
    if (t === "s") {
      const idx = parseInt(rawV, 10);
      rawV = sstArray?.[idx] ?? "";
      typ = "s";
    }
    if (!vNode && fNode) {
      rawV = fNode.textContent;
      typ = "str";
    }
    const { r, c: cc } = cellRefToRC(ref);
    const cell = formatCell(rawV, typ, isNaN(sIdx) ? -1 : sIdx, styles);
    const rr = r - range.r0, rc = cc - range.c0;
    if (rows[rr]) rows[rr][rc] = cell;
  }
  return { rows, range };
}

/* =================== NOMBRE HOJA Y TABLAS =================== */
const SHEET_FORMULA = "FORMULAS"; // hoja con tabla Formula
const TABLE_FORMULA = "formula";
const SHEET_RACIONES = "raciones"; // hoja con tabla Comida
const TABLE_COMIDA = "comida";

/* =================== LECTURAS =================== */
async function readXlsxTableFormula(file) {
  const zip = await JSZip.loadAsync(file);
  const wbXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!wbXml) throw new Error("XLSX inválido: falta workbook.xml");
  const wbDoc = xmlToDoc(wbXml);
  const sheets = Array.from(wbDoc.getElementsByTagName("sheet")).map((s) => ({
    name: s.getAttribute("name"),
    rId: s.getAttribute("r:id"),
  }));
  const targetSheet = sheets.find((s) => norm(s.name) === norm(SHEET_FORMULA));
  if (!targetSheet)
    throw new Error(`No existe una hoja llamada '${SHEET_FORMULA}'.`);

  const wbRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  const wbRelsDoc = xmlToDoc(wbRelsXml);
  const wbRels = Object.fromEntries(
    Array.from(wbRelsDoc.getElementsByTagName("Relationship")).map((r) => [
      r.getAttribute("Id"),
      r.getAttribute("Target"),
    ])
  );
  const sheetRelTarget = wbRels[targetSheet.rId];
  const sheetPath = `xl/${sheetRelTarget}`;
  const sheetXml = await zip.file(sheetPath)?.async("string");
  if (!sheetXml) throw new Error(`No pude abrir la hoja '${SHEET_FORMULA}'.`);

  // SharedStrings + estilos
  let sstArray = [];
  if (zip.file("xl/sharedStrings.xml")) {
    const sstXml = await zip.file("xl/sharedStrings.xml").async("string");
    const sstDoc = xmlToDoc(sstXml);
    sstArray = Array.from(sstDoc.getElementsByTagName("si")).map((si) => {
      const ts = si.getElementsByTagName("t");
      let txt = "";
      for (let i = 0; i < ts.length; i++) txt += ts[i].textContent;
      return txt;
    });
  }
  let styles = { xfToFmt: {}, numFmtIdToCode: {} };
  if (zip.file("xl/styles.xml")) {
    styles = parseStyles(await zip.file("xl/styles.xml").async("string"));
  }

  // Buscar tabla 'Formula' en relaciones de hoja
  const sheetRelsPath = `xl/worksheets/_rels/${sheetRelTarget.split("/").pop()}.rels`;
  const sheetRelsXml = await zip.file(sheetRelsPath)?.async("string");
  let tableMeta = null;
  if (sheetRelsXml) {
    const sRelsDoc = xmlToDoc(sheetRelsXml);
    const wsDoc = xmlToDoc(sheetXml);
    const ids = Array.from(wsDoc.getElementsByTagName("tablePart")).map((n) =>
      n.getAttribute("r:id")
    );
    const relNodes = Array.from(
      sRelsDoc.getElementsByTagName("Relationship")
    );
    for (const rid of ids) {
      const rel = relNodes.find((n) => n.getAttribute("Id") === rid);
      if (!rel) continue;
      const tgt = rel.getAttribute("Target");
      const fullPath = (function () {
        const baseParts = ("xl/" + sheetRelTarget).split("/");
        baseParts.pop();
        let r = tgt.replace(/^\.\//, "");
        while (r.startsWith("../")) {
          r = r.slice(3);
          if (baseParts.length > 1) baseParts.pop();
        }
        return baseParts.concat(r.split("/")).join("/");
      })();
      const tXml = await zip.file(fullPath)?.async("string");
      if (!tXml) continue;
      const tDoc = xmlToDoc(tXml);
      const tNode = tDoc.getElementsByTagName("table")[0];
      const disp = tNode?.getAttribute("displayName");
      const name = tNode?.getAttribute("name");
      if (
        (disp && norm(disp) === TABLE_FORMULA) ||
        (name && norm(name) === TABLE_FORMULA)
      ) {
        tableMeta = { path: fullPath, ref: tNode.getAttribute("ref") };
        break;
      }
    }
  }
  if (!tableMeta) {
    const tableFiles = Object.keys(zip.files).filter((p) =>
      /^xl\/tables\/.+\.xml$/i.test(p)
    );
    for (const p of tableFiles) {
      const tXml = await zip.file(p).async("string");
      const tDoc = xmlToDoc(tXml);
      const tNode = tDoc.getElementsByTagName("table")[0];
      const disp = tNode?.getAttribute("displayName");
      const name = tNode?.getAttribute("name");
      if (norm(disp || name) === TABLE_FORMULA) {
        tableMeta = { path: p, ref: tNode.getAttribute("ref") };
        break;
      }
    }
  }
  if (!tableMeta)
    throw new Error(
      `No encontré una tabla llamada 'Formula' en la hoja '${SHEET_FORMULA}'.`
    );

  const range = parseRangeRef(tableMeta.ref);
  if (!range) throw new Error("La tabla 'Formula' no tiene un rango válido.");

  const full = parseSheetXml(sheetXml, sstArray, styles);
  const baseR = full.range?.r0 ?? 0;
  const baseC = full.range?.c0 ?? 0;
  const rows = [];
  for (let r = range.r0; r <= range.r1; r++) {
    const arr = [];
    for (let c = range.c0; c <= range.c1; c++) {
      const rr = r - baseR, cc = c - baseC;
      const cell = full.rows[rr]?.[cc] ?? { display: "", kind: "empty" };
      arr.push(cell);
    }
    rows.push(arr);
  }
  return { rows };
}

/* === LEE ORDEN DE DESCARGA: tabla DescargaCorrales (si existe) === */
async function readMixerOrderFromExcel(file) {
  const zip = await JSZip.loadAsync(file);
  const wbXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!wbXml) return null;
  const wbDoc = xmlToDoc(wbXml);

  const sheets = Array.from(wbDoc.getElementsByTagName("sheet")).map((s) => ({
    name: s.getAttribute("name"),
    rId: s.getAttribute("r:id"),
  }));

  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  const relsDoc = xmlToDoc(relsXml);
  const wbRels = Object.fromEntries(
    Array.from(relsDoc.getElementsByTagName("Relationship")).map((r) => [
      r.getAttribute("Id"),
      r.getAttribute("Target"),
    ])
  );

  let tableMeta = null;
  let sheetRelTarget = null;

  for (const sh of sheets) {
    const target = wbRels[sh.rId];
    const sheetPath = `xl/${target}`;
    const sheetXml = await zip.file(sheetPath)?.async("string");
    if (!sheetXml) continue;

    const sheetRelsPath = `xl/worksheets/_rels/${target.split("/").pop()}.rels`;
    const sheetRelsXml = await zip.file(sheetRelsPath)?.async("string");
    if (!sheetRelsXml) continue;

    const wsDoc = xmlToDoc(sheetXml);
    const sRelsDoc = xmlToDoc(sheetRelsXml);
    const partIds = Array.from(wsDoc.getElementsByTagName("tablePart")).map(
      (n) => n.getAttribute("r:id")
    );
    const relNodes = Array.from(
      sRelsDoc.getElementsByTagName("Relationship")
    );

    for (const rid of partIds) {
      const rel = relNodes.find((n) => n.getAttribute("Id") === rid);
      if (!rel) continue;
      const tgt = rel.getAttribute("Target");
      const fullPath = (function () {
        const baseParts = ("xl/" + target).split("/");
        baseParts.pop();
        let r = tgt.replace(/^\.\//, "");
        while (r.startsWith("../")) {
          r = r.slice(3);
          if (baseParts.length > 1) baseParts.pop();
        }
        return baseParts.concat(r.split("/")).join("/");
      })();
      const tXml = await zip.file(fullPath)?.async("string");
      if (!tXml) continue;
      const tDoc = xmlToDoc(tXml);
      const tNode = tDoc.getElementsByTagName("table")[0];
      const nm = (tNode?.getAttribute("displayName") ||
        tNode?.getAttribute("name") ||
        ""
      )
        .toLowerCase()
        .replace(/\s+/g, "");
      if (nm === "descargacorrales") {
        tableMeta = { path: fullPath, ref: tNode.getAttribute("ref") };
        sheetRelTarget = target;
        break;
      }
    }
    if (tableMeta) break;
  }
  if (!tableMeta) return null;

  const pageXml = await zip.file(`xl/${sheetRelTarget}`)?.async("string");
  const sstXml = await zip.file("xl/sharedStrings.xml")?.async("string");

  let sstArray = [];
  if (sstXml) {
    const sstDoc = xmlToDoc(sstXml);
    sstArray = Array.from(sstDoc.getElementsByTagName("si")).map((si) => {
      const ts = si.getElementsByTagName("t");
      let txt = "";
      for (let i = 0; i < ts.length; i++) txt += ts[i].textContent;
      return txt;
    });
  }
  let styles = { xfToFmt: {}, numFmtIdToCode: {} };
  if (zip.file("xl/styles.xml")) {
    styles = parseStyles(await zip.file("xl/styles.xml").async("string"));
  }

  const range = parseRangeRef(tableMeta.ref);
  if (!range) return null;

  const full = parseSheetXml(pageXml, sstArray, styles);
  const baseR = full.range?.r0 ?? 0;
  const baseC = full.range?.c0 ?? 0;

  const tableRows = [];
  for (let r = range.r0; r <= range.r1; r++) {
    const arr = [];
    for (let c = range.c0; c <= range.c1; c++) {
      const rr = r - baseR, cc = c - baseC;
      arr.push(full.rows[rr]?.[cc] ?? { display: "", kind: "empty" });
    }
    tableRows.push(arr);
  }

  const header = tableRows[0]?.map((c) => c.display || "") || [];
  const findCol = (label) => {
    const t = (label || "").toLowerCase().replace(/\s+/g, "");
    return header.findIndex(
      (h) => (h || "").toLowerCase().replace(/\s+/g, "") === t
    );
  };
  const idxM = [1, 2, 3, 4, 5].map((n) => findCol(`mixer${n}`));
  const readList = (ix) =>
    ix < 0
      ? []
      : tableRows
          .slice(1)
          .map((r) => r[ix]?.display ?? "")
          .filter((x) => String(x).trim() !== "");
  const order = {};
  idxM.forEach((ix, i) => {
    if (ix >= 0) order[String(i + 1)] = readList(ix);
  });
  return order;
}

async function readXlsxTableComidaSelected(file) {
  const zip = await JSZip.loadAsync(file);
  const wbXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!wbXml) throw new Error("XLSX inválido: falta workbook.xml");
  const wbDoc = xmlToDoc(wbXml);
  const sheets = Array.from(wbDoc.getElementsByTagName("sheet")).map((s) => ({
    name: s.getAttribute("name"),
    rId: s.getAttribute("r:id"),
  }));
  const targetSheet = sheets.find((s) => norm(s.name) === norm(SHEET_RACIONES));
  if (!targetSheet)
    throw new Error(`No existe una hoja llamada '${SHEET_RACIONES}'.`);

  const wbRelsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  const wbRelsDoc = xmlToDoc(wbRelsXml);
  const wbRels = Object.fromEntries(
    Array.from(wbRelsDoc.getElementsByTagName("Relationship")).map((r) => [
      r.getAttribute("Id"),
      r.getAttribute("Target"),
    ])
  );
  const sheetRelTarget = wbRels[targetSheet.rId];
  const sheetPath = `xl/${sheetRelTarget}`;
  const sheetXml = await zip.file(sheetPath)?.async("string");
  if (!sheetXml) throw new Error(`No pude abrir la hoja '${SHEET_RACIONES}'.`);

  // SharedStrings + estilos
  let sstArray = [];
  if (zip.file("xl/sharedStrings.xml")) {
    const sstXml = await zip.file("xl/sharedStrings.xml").async("string");
    const sstDoc = xmlToDoc(sstXml);
    sstArray = Array.from(sstDoc.getElementsByTagName("si")).map((si) => {
      const ts = si.getElementsByTagName("t");
      let txt = "";
      for (let i = 0; i < ts.length; i++) txt += ts[i].textContent;
      return txt;
    });
  }
  let styles = { xfToFmt: {}, numFmtIdToCode: {} };
  if (zip.file("xl/styles.xml")) {
    styles = parseStyles(await zip.file("xl/styles.xml").async("string"));
  }

  // localizar tabla 'Comida' por relaciones
  const sheetRelsPath = `xl/worksheets/_rels/${sheetRelTarget
    .split("/")
    .pop()}.rels`;
  const sheetRelsXml = await zip.file(sheetRelsPath)?.async("string");
  let tableMeta = null;
  if (sheetRelsXml) {
    const sRelsDoc = xmlToDoc(sheetRelsXml);
    const wsDoc = xmlToDoc(sheetXml);
    const ids = Array.from(wsDoc.getElementsByTagName("tablePart")).map((n) =>
      n.getAttribute("r:id")
    );
    const relNodes = Array.from(
      sRelsDoc.getElementsByTagName("Relationship")
    );
    for (const rid of ids) {
      const rel = relNodes.find((n) => n.getAttribute("Id") === rid);
      if (!rel) continue;
      const tgt = rel.getAttribute("Target");
      const fullPath = (function () {
        const baseParts = ("xl/" + sheetRelTarget).split("/");
        baseParts.pop();
        let r = tgt.replace(/^\.\//, "");
        while (r.startsWith("../")) {
          r = r.slice(3);
          if (baseParts.length > 1) baseParts.pop();
        }
        return baseParts.concat(r.split("/")).join("/");
      })();
      const tXml = await zip.file(fullPath)?.async("string");
      if (!tXml) continue;
      const tDoc = xmlToDoc(tXml);
      const tNode = tDoc.getElementsByTagName("table")[0];
      const disp = tNode?.getAttribute("displayName");
      const name = tNode?.getAttribute("name");
      if (
        (disp && norm(disp) === TABLE_COMIDA) ||
        (name && norm(name) === TABLE_COMIDA)
      ) {
        tableMeta = { path: fullPath, ref: tNode.getAttribute("ref") };
        break;
      }
    }
  }
  if (!tableMeta) {
    const tableFiles = Object.keys(zip.files).filter((p) =>
      /^xl\/tables\/.+\.xml$/i.test(p)
    );
    for (const p of tableFiles) {
      const tXml = await zip.file(p).async("string");
      const tDoc = xmlToDoc(tXml);
      const tNode = tDoc.getElementsByTagName("table")[0];
      const disp = tNode?.getAttribute("displayName");
      const name = tNode?.getAttribute("name");
      if (norm(disp || name) === TABLE_COMIDA) {
        tableMeta = { path: p, ref: tNode.getAttribute("ref") };
        break;
      }
    }
  }
  if (!tableMeta) throw new Error("No encontré la tabla 'Comida' en 'raciones'.");

  const range = parseRangeRef(tableMeta.ref);
  if (!range) throw new Error("La tabla 'Comida' no tiene un rango válido.");

  const full = parseSheetXml(sheetXml, sstArray, styles);
  const baseR = full.range?.r0 ?? 0;
  const baseC = full.range?.c0 ?? 0;

  // Extraer toda la tabla
  const rawRows = [];
  for (let r = range.r0; r <= range.r1; r++) {
    const arr = [];
    for (let c = range.c0; c <= range.c1; c++) {
      const rr = r - baseR, cc = c - baseC;
      const cell = full.rows[rr]?.[cc] ?? { display: "", kind: "empty" };
      arr.push(cell);
    }
    rawRows.push(arr);
  }

  // Quedarme solo con: Corral, Cab, Kg, Etapa, KgTc/Dia, Mixer, Ajuste (en ese orden)
  const wantedOrder = ["Corral","Cab","Kg","Etapa","KgTc/Dia","Mixer","Ajuste"];
  const header = rawRows[0]?.map((c) => c.display || "") || [];
  const idxMap = wantedOrder.map((w) =>
    header.findIndex((h) => norm(h) === norm(w))
  );

  const rows = [wantedOrder.map((w) => ({ display: w, kind: "text" }))];
  for (let i = 1; i < rawRows.length; i++) {
    const src = rawRows[i];
    const out = idxMap.map((ix) =>
      ix >= 0 ? (src[ix] ?? { display: "", kind: "empty" }) : { display: "", kind: "empty" }
    );
    rows.push(out);
  }
  return { rows };
}

/* =================== COMPONENTE =================== */
export default function App() {
  const [rowsFormula, setRowsFormula] = useState(null);
  const [rowsRacion, setRowsRacion] = useState(null);
  const [err, setErr] = useState("");
  const [tab, setTab] = useState("Formula");

  // Orden editable de corrales por mixer + drag
  const [mixerOrder, setMixerOrder] = useState({});
  const [dragInfo, setDragInfo] = useState(null);
  function onDragStartMixer(mixer, fromIndex) {
    setDragInfo({ mixer, fromIndex });
  }
  function onDropMixer(mixer, toIndex, currentList) {
    setMixerOrder((prev) => {
      const base =
        prev[mixer] && prev[mixer].length
          ? [...prev[mixer]]
          : currentList.map((r) => r.corral);
      const arr = [...base];
      const from = dragInfo?.fromIndex ?? toIndex;
      const [moved] = arr.splice(from, 1);
      arr.splice(toIndex, 0, moved);
      return { ...prev, [mixer]: arr };
    });
    setDragInfo(null);
  }

  async function onFile(e) {
    setErr("");
    setRowsFormula(null);
    setRowsRacion(null);
    try {
      const f = e.target.files?.[0];
      if (!f) return;
      const buf = await f.arrayBuffer();
      const [{ rows: rowsF }, { rows: rowsC }] = await Promise.all([
        readXlsxTableFormula(buf),
        readXlsxTableComidaSelected(buf),
      ]);
      setRowsFormula(rowsF);
      setRowsRacion(rowsC);

      // Orden inicial leído de "DescargaCorrales" (si existe)
      try {
        const order = await readMixerOrderFromExcel(buf);
        if (order && Object.keys(order).length) setMixerOrder(order);
      } catch {/* noop */}
    } catch (ex) {
      setErr(ex?.message || "Error al leer");
    }
  }

  // ===== Helpers Corrales/Racion =====
  const headerR = rowsRacion?.[0]?.map((c) => c.display || "") || [];
  // FIX: comparar contra lower-case (porque norm() devuelve lower-case)
  const idxEtapa  = headerR.findIndex((h) => norm(h) === "etapa");
  const idxKgTC   = headerR.findIndex((h) => norm(h) === "kgtc/dia");
  const idxMixer  = headerR.findIndex((h) => norm(h) === "mixer");
  const idxAjuste = headerR.findIndex((h) => norm(h) === "ajuste");
  const idxCab    = headerR.findIndex((h) => norm(h) === "cab");
  const idxCorral = headerR.findIndex((h) => norm(h) === "corral");

  // Opciones Etapa desde encabezados de Formula
  const headerF = rowsFormula?.[0]?.map((c) => c.display || "") || [];
  const etapaOptions = headerF
    .map((x) => String(x))
    .filter((x) => x && norm(x) !== "etapa" && norm(x) !== "insumos" && !/%/.test(x));

  // Edición de celdas
  function updateRacionCell(rIndex, cIndex, newDisplay) {
    setRowsRacion((prev) => {
      if (!prev) return prev;
      const next = prev.map((row, ri) =>
        ri === rIndex
          ? row.map((cell, ci) =>
              ci === cIndex ? { ...cell, display: newDisplay } : cell
            )
          : row
      );
      return next;
    });
  }

  function normalizePercentInput(v) {
    if (v == null) return "";
    let s = String(v).trim();
    s = s.replace(",", ".");
    if (s.endsWith("%")) s = s.slice(0, -1);
    if (s === "") return "";
    const n = Number(s);
    if (Number.isNaN(n)) return "";
    return `${n}%`;
  }

  function parseNum(x) {
    const s = String(x ?? "").replace(",", ".");
    const n = Number(s.replace(/[^0-9.+-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  }
  function percentToFactor(x) {
    const s = String(x ?? "").trim().replace(",", ".");
    if (s === "") return 1;
    if (s.endsWith("%")) {
      const n = Number(s.slice(0, -1));
      return Number.isFinite(n) ? n / 100 : 1;
    }
    const n = Number(s);
    if (!Number.isFinite(n)) return 1;
    return n > 1 ? n / 100 : n; // admite 50 o 0.5
  }

  // Kg/Ronda por fila
  function computeKgRondaRow(row) {
    const cab = idxCab >= 0 ? parseNum(row[idxCab]?.display) : 0;
    const kgtc = idxKgTC >= 0 ? parseNum(row[idxKgTC]?.display) : 0;
    const adj = idxAjuste >= 0 ? percentToFactor(row[idxAjuste]?.display) : 1;
    return (cab * kgtc * adj) / 2;
  }

  // === Helpers Mixer: insumos/factores según etapa (desde Formula) ===
  function getEtapaColIndex(etapa) {
    if (!rowsFormula) return -1;
    const hdr = rowsFormula[0]?.map((c) => c.display || "") || [];
    return hdr.findIndex((h) => norm(h) === norm(etapa));
  }
  function getInsumosFactorsByEtapa(etapa) {
    const col = getEtapaColIndex(etapa);
    if (col < 0) return [];
    const out = [];
    const ban = new Set(["%ms", "%mv", "total", "codigo", "dieta"]);
    for (let i = 1; i < rowsFormula.length; i++) {
      const ins = rowsFormula[i]?.[0]?.display ?? "";
      if (!ins) continue;
      const insNorm = norm(String(ins));
      if (ban.has(insNorm)) continue;
      const val = rowsFormula[i]?.[col]?.display ?? "";
      if (val === "") continue;
      const factor = percentToFactor(val);
      if (factor > 0) out.push({ insumo: String(ins), factor });
    }
    return out;
  }

  /* =================== mixerData (se recalcula con drag) =================== */
  const mixerData = React.useMemo(() => {
    if (!rowsRacion) return null;
    const result = { byMixer: {} };
    const mixers = ["1", "2", "3", "4", "5"];

    for (const m of mixers) {
      // Filas base del mixer
      const filasBase = rowsRacion
        .slice(1)
        .filter((row) => (row[idxMixer]?.display || "").toString() === m);

      // Ordenar según mixerOrder actual
      const orden = mixerOrder[m] || [];
      const filas = [...filasBase].sort(
        (a, b) =>
          orden.indexOf(a[idxCorral]?.display) -
          orden.indexOf(b[idxCorral]?.display)
      );

      // Etapa, receta y total
      const etapa =
        filas.find((r) => (r[idxEtapa]?.display || "") !== "")?.[idxEtapa]
          ?.display || "";
      const receta = etapa ? getInsumosFactorsByEtapa(etapa) : [];
      const total = filas.reduce((acc, row) => acc + computeKgRondaRow(row), 0);

      // Carga acumulada
      let acum = 0;
      const carga = receta.map(({ insumo, factor }) => {
        const kg = total * factor;
        acum += kg;
        return { insumo, kg, balanza: acum };
      });

      // Descarga según orden → la última balanza queda en 0
      let restante = total;
      const descarga = filas.map((row) => {
        const corral = row[idxCorral]?.display ?? "";
        const kg = computeKgRondaRow(row);
        restante -= kg;
        if (restante < 0) restante = 0;
        return { corral, kg, balanza: restante };
      });

      result.byMixer[m] = { total, etapa, carga, descarga };
    }
    return result;
  }, [rowsRacion, rowsFormula, mixerOrder]);

  /* =================== RENDER =================== */
  return (
    <div className="p-4 space-y-4">
      <input type="file" accept=".xlsx" onChange={onFile} />
      {err && (
        <div className="bg-red-50 border border-red-300 text-red-800 p-2 rounded">
          {err}
        </div>
      )}

      {/* Tabs */}
      <div className="flex border-b gap-2">
        <button
          onClick={() => setTab("Formula")}
          className={`px-3 py-2 -mb-px border-b-2 ${tab === "Formula" ? "border-black font-semibold" : "border-transparent"}`}
        >
          Formula
        </button>
        <button
          onClick={() => setTab("Corrales/Racion")}
          className={`px-3 py-2 -mb-px border-b-2 ${tab === "Corrales/Racion" ? "border-black font-semibold" : "border-transparent"}`}
        >
          Corrales/Racion
        </button>
        <button
          onClick={() => setTab("Mixer")}
          className={`px-3 py-2 -mb-px border-b-2 ${tab === "Mixer" ? "border-black font-semibold" : "border-transparent"}`}
        >
          Mixer
        </button>
      </div>

      {/* Formula */}
      {tab === "Formula" && rowsFormula && (
        <div className="table-wrap">
          <table className="min-w-full border-collapse">
            <tbody>
              {rowsFormula.map((r, ri) => (
                <tr key={ri}>
                  {r.map((cell, ci) => {
                    const isHeader = ri === 0;
                    const content = isHeader ? prettyHeader(cell.display) : cell.display;
                    const baseCls = `px-2 py-1 text-sm ${isHeader ? "font-semibold" : ""} ${!isHeader && (cell.kind==="number"||cell.kind==="percent") ? "text-right" : "text-left"}`;
                    return isHeader ? (
                      <th
                        key={ci}
                        className={baseCls}
                        style={{ borderBottom: "1px solid #e5f3ec", background: "rgba(16,185,129,.08)" }}
                      >
                        {content}
                      </th>
                    ) : (
                      <td
                        key={ci}
                        className={baseCls}
                        style={{ borderBottom: "1px solid #f0faf6" }}
                      >
                        {content}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Corrales/Racion */}
      {tab === "Corrales/Racion" && rowsRacion && (
        <div className="table-wrap">
          <table className="min-w-full border-collapse">
            <tbody>
              {rowsRacion.map((r, ri) => (
                <tr key={ri}>
                  {r.map((cell, ci) => {
                    // Encabezados con mayúscula + negrita
                    if (ri === 0) {
                      return (
                        <th
                          key={ci}
                          className="px-2 py-1 text-sm font-semibold text-left"
                          style={{ borderBottom: "1px solid #e5f3ec", background: "rgba(16,185,129,.08)" }}
                        >
                          {prettyHeader(cell.display)}
                        </th>
                      );
                    }

                    // === Celda ETAPA con color solo si hay valor ===
                    if (ci === idxEtapa && idxEtapa >= 0) {
                      const etapaVal = cell.display || "";
                      const etapaClass = getEtapaClass(etapaVal); // "" si está vacío
                      return (
                        <td key={ci} className="border px-2 py-1 text-sm">
                          <div className={`etapa-pill ${etapaClass}`}>
                            <select
                              className="etapa-select"
                              value={etapaVal}
                              onChange={(e) => updateRacionCell(ri, ci, e.target.value)}
                            >
                              <option value=""> </option>
                              {etapaOptions.map((opt) => (
                                <option key={opt} value={opt}>{opt}</option>
                              ))}
                            </select>
                          </div>
                        </td>
                      );
                    }

                    // === KgTC/Dia editable ===
                    if (ci === idxKgTC && idxKgTC >= 0) {
                      const valNum = (() => {
                        const s = String(cell.display || "")
                          .replace(",", " .")
                          .replace(" ", "")
                          .replace(" .", ".");
                        const n = Number(s);
                        return Number.isFinite(n) ? n : "";
                      })();
                      return (
                        <td key={ci} className="border px-2 py-1 text-sm text-right">
                          <input
                            type="number"
                            step="0.5"
                            className="w-24 text-right outline-none"
                            value={valNum}
                            onChange={(e) => updateRacionCell(ri, ci, e.target.value)}
                          />
                        </td>
                      );
                    }

                    // === Mixer editable ===
                    if (ci === idxMixer && idxMixer >= 0) {
                      const options = ["sin mixer", "1", "2", "3", "4", "5"];
                      const current = (cell.display || "").toString();
                      return (
                        <td key={ci} className="border px-2 py-1 text-sm">
                          <select
                            className="outline-none w-full"
                            value={current}
                            onChange={(e) => updateRacionCell(ri, ci, e.target.value)}
                          >
                            {options.map((opt) => (
                              <option key={opt} value={opt}>{opt}</option>
                            ))}
                          </select>
                        </td>
                      );
                    }

                    // === Ajuste % editable ===
                    if (ci === idxAjuste && idxAjuste >= 0) {
                      const shown = String(cell.display || "");
                      return (
                        <td key={ci} className="border px-2 py-1 text-sm text-right">
                          <input
                            type="text"
                            className="w-20 text-right outline-none"
                            value={shown}
                            onChange={(e) => updateRacionCell(ri, ci, e.target.value)}
                            onBlur={(e) =>
                              updateRacionCell(ri, ci, normalizePercentInput(e.target.value))
                            }
                          />
                        </td>
                      );
                    }

                    // por defecto
                    return (
                      <td
                        key={ci}
                        className={`border px-2 py-1 text-sm ${
                          cell.kind === "number" || cell.kind === "percent" ? "text-right" : "text-left"
                        }`}
                      >
                        {cell.display}
                      </td>
                    );
                  })}

                  {/* Kg/Ronda */}
                  {ri === 0 ? (
                    <th className="px-2 py-1 text-sm font-semibold text-right" style={{ borderBottom: "1px solid #e5f3ec", background: "rgba(16,185,129,.08)" }}>
                      Kg/Ronda
                    </th>
                  ) : (
                    (() => {
                      const val = computeKgRondaRow(r);
                      const shown = Number.isFinite(val) ? val.toFixed(2) : "";
                      return <td className="border px-2 py-1 text-sm text-right">{shown}</td>;
                    })()
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Mixer */}
      {tab === "Mixer" && mixerData && (
        <div className="space-y-8">
          {Object.entries(mixerData.byMixer).map(([m, data]) => (
            <MixerBox key={m} etapa={data.etapa} titulo={data.etapa || "Inicio"}>
              <div className="grid md:grid-cols-2 gap-6">
                {/* CARGA PALA */}
                <div>
                  <div className="font-semibold mb-1">
                    MIXER {m} — CargaPala {data.etapa ? `(Etapa: ${data.etapa})` : ""}
                  </div>
                  <table className="min-w-full border-collapse">
                    <tbody>
                      <tr>
                        <td className="border px-2 py-1 text-sm font-semibold">INSUMOS</td>
                        <td className="border px-2 py-1 text-sm font-semibold text-right">KG/ INSUMO</td>
                        <td className="border px-2 py-1 text-sm font-semibold text-right">BALANZA</td>
                      </tr>
                      {data.carga.length === 0 ? (
                        <tr>
                          <td className="border px-2 py-1 text-sm" colSpan={3}>Sin datos</td>
                        </tr>
                      ) : (
                        data.carga.map((row, i) => (
                          <tr key={i}>
                            <td className="border px-2 py-1 text-sm">{row.insumo}</td>
                            <td className="border px-2 py-1 text-sm text-right">{row.kg.toFixed(0)} kg</td>
                            <td className="border px-2 py-1 text-sm text-right">{row.balanza.toFixed(0)} kg</td>
                          </tr>
                        ))
                      )}
                      <tr>
                        <td className="border px-2 py-1 text-sm font-semibold">CARGA TOTAL</td>
                        <td className="border px-2 py-1 text-sm font-semibold text-right">{data.total.toFixed(0)} kg</td>
                        <td className="border px-2 py-1 text-sm font-semibold text-right">{data.total.toFixed(0)} kg</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* DESCARGA CORRALES */}
                <div>
                  <div className="font-semibold mb-1">MIXER {m} — DescargaCorrales</div>
                  <table className="min-w-full border-collapse">
                    <tbody>
                      <tr>
                        <td className="border px-2 py-1 text-sm font-semibold">Corral</td>
                        <td className="border px-2 py-1 text-sm font-semibold text-right">KG</td>
                        <td className="border px-2 py-1 text-sm font-semibold text-right">BALANZA</td>
                      </tr>
                      {data.descarga.length === 0 ? (
                        <tr>
                          <td className="border px-2 py-1 text-sm" colSpan={3}>Sin corrales asignados</td>
                        </tr>
                      ) : (
                        data.descarga.map((row, i) => (
                          <tr
                            key={`${row.corral}-${i}`}
                            draggable
                            onDragStart={() => onDragStartMixer(m, i)}
                            onDragOver={(e) => e.preventDefault()}
                            onDrop={() => onDropMixer(m, i, data.descarga)}
                          >
                            <td className="border px-2 py-1 text-sm">{row.corral}</td>
                            <td className="border px-2 py-1 text-sm text-right">{row.kg.toFixed(0)} kg</td>
                            <td className="border px-2 py-1 text-sm text-right">{row.balanza.toFixed(0)} kg</td>
                          </tr>
                        ))
                      )}
                      <tr>
                        <td className="border px-2 py-1 text-sm font-semibold">TOTAL</td>
                        <td className="border px-2 py-1 text-sm font-semibold text-right">{data.total.toFixed(0)} kg</td>
                        <td className="border px-2 py-1 text-sm font-semibold text-right">{data.total.toFixed(0)} kg</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            </MixerBox>
          ))}
        </div>
      )}
    </div>
  );
}
