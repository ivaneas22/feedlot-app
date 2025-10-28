import React, { useState } from "react";
import JSZip from "jszip";
// --- Paso 2: Cálculos de kilos por corral y mixer ---

function calcularMixers(sheetData) {
  // Suponemos que la hoja tiene columnas:
  // Corral | Cabezas | Dieta | Kg x cab | Ajuste (%) | Mixer
  // (podés adaptar los nombres según tu Excel real)
  const mixers = {};

  sheetData.forEach((fila) => {
    const [corral, cabezas, dieta, kgxCab, ajuste, mixer] = fila;

    const cab = Number(cabezas) || 0;
    const kg = Number(kgxCab) || 0;
    const adj = Number(ajuste) || 100;
    const mix = mixer || "Sin asignar";

    const kgTotales = cab * kg * (adj / 100);
    const kgPorRonda = kgTotales / 2;

    if (!mixers[mix]) {
      mixers[mix] = {
        mixer: mix,
        total: 0,
        corrales: [],
      };
    }

    mixers[mix].total += kgTotales;
    mixers[mix].corrales.push({
      corral,
      dieta,
      cabezas: cab,
      kgxCab: kg,
      ajuste: adj,
      kgTotales,
      kgPorRonda,
    });
  });

  return Object.values(mixers);
}

export default function App() {
  const [fileName, setFileName] = useState("");
  const [sheets, setSheets] = useState([]);
  const [error, setError] = useState("");

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setError("");
      setFileName(file.name);
      const data = await file.arrayBuffer();
      const zip = await JSZip.loadAsync(data);

      // Archivos XML principales del XLSX
      const workbook = await zip.file("xl/workbook.xml").async("string");
      const rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");

      // Extrae nombres de hojas
      const sheetNames = [...workbook.matchAll(/<sheet name="([^"]+)"/g)].map(
        (m) => m[1]
      );
      const sheetPaths = [...rels.matchAll(/Target="worksheets\/(.*?)"/g)].map(
        (m) => m[1]
      );

      const allSheets = [];
      for (let i = 0; i < sheetNames.length; i++) {
        const xml = await zip
          .file(`xl/worksheets/${sheetPaths[i]}`)
          .async("string");

        const rows = [...xml.matchAll(/<row.*?<\/row>/g)].map((r) =>
          [...r[0].matchAll(/<v.*?>(.*?)<\/v>/g)].map((c) => c[1])
        );
        allSheets.push({ name: sheetNames[i], data: rows });
      }
      setSheets(allSheets);
    } catch (err) {
      console.error(err);
      setError("Error leyendo el archivo. Asegurate que sea un .xlsx válido.");
    }
  }

  return (
    <div style={{ padding: 20, fontFamily: "Poppins, sans-serif" }}>
      <h1>📘 Lector de Excel (Feedlot)</h1>

      <input
        type="file"
        accept=".xlsx,.xlsm"
        onChange={handleFile}
        style={{
          padding: 8,
          border: "1px solid #ccc",
          borderRadius: 8,
          marginBottom: 10,
        }}
      />

      {fileName && <p><b>Archivo cargado:</b> {fileName}</p>}
      {error && <p style={{ color: "red" }}>{error}</p>}

      {sheets.map((sheet, i) => (
        <div
          key={i}
          style={{
            marginTop: 20,
            background: "#f9fafb",
            border: "1px solid #e2e8f0",
            borderRadius: 10,
            padding: 10,
          }}
        >
          <h3>{sheet.name}</h3>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <tbody>
              {sheet.data.slice(0, 10).map((row, rIdx) => (
                <tr key={rIdx}>
                  {row.map((cell, cIdx) => (
                    <td
                      key={cIdx}
                      style={{
                        border: "1px solid #ddd",
                        padding: "4px 8px",
                        fontSize: "0.9em",
                      }}
                    >
                      {cell}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
      {/* ======== PANEL DE MIXERS (PASO 3) ======== */}
{(() => {
  // Tomamos la hoja llamada "Comida" o la primera que contenga datos
  const hojaComida = sheets.find((s) =>
    s.name.toLowerCase().includes("comida")
  );

  if (!hojaComida || hojaComida.data.length < 2) return null;

  // Quitamos encabezado
  const filas = hojaComida.data.slice(1);
  const mixers = calcularMixers(filas);

  return (
    <div style={{ marginTop: 30 }}>
      <h2>🚜 Mixers calculados</h2>
      {mixers.map((mix) => (
        <div
          key={mix.mixer}
          style={{
            marginTop: 20,
            background: "#ffffff",
            border: "2px solid #10b98130",
            borderRadius: 12,
            boxShadow: "0 2px 10px rgba(0,0,0,0.05)",
            padding: 15,
          }}
        >
          <h3 style={{ color: "#047857", marginBottom: 10 }}>
            Mixer {mix.mixer}
          </h3>
          <p>
            <b>Total:</b> {mix.total.toLocaleString("es-AR")} kg
          </p>

          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr
                style={{
                  background: "#ecfdf5",
                  borderBottom: "1px solid #d1fae5",
                }}
              >
                <th>Corral</th>
                <th>Dieta</th>
                <th>Cabezas</th>
                <th>Kg/cab</th>
                <th>Ajuste %</th>
                <th>Kg totales</th>
                <th>Kg ronda</th>
              </tr>
            </thead>
            <tbody>
              {mix.corrales.map((c, i) => {
                const color =
                  c.dieta.toLowerCase().includes("termin")
                    ? "#fee2e2"
                    : c.dieta.toLowerCase().includes("recr")
                    ? "#dbeafe"
                    : "#fef9c3";
                return (
                  <tr
                    key={i}
                    style={{
                      background: color,
                      borderBottom: "1px solid #f1f5f9",
                    }}
                  >
                    <td>{c.corral}</td>
                    <td>{c.dieta}</td>
                    <td>{c.cabezas}</td>
                    <td>{c.kgxCab}</td>
                    <td>{c.ajuste}</td>
                    <td>{c.kgTotales.toFixed(0)}</td>
                    <td>{c.kgPorRonda.toFixed(0)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
})()}

    </div>
  );
}
