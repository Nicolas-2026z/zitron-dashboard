import { useState } from "react";

const CLAVE = "zitron2026!";

function Login({ onAcceso }) {
  const [input, setInput] = useState("");
  const [error, setError] = useState(false);
  const [show, setShow] = useState(false);

  const intentar = () => {
    if (input === CLAVE) {
      onAcceso();
    } else {
      setError(true);
      setInput("");
      setTimeout(() => setError(false), 2000);
    }
  };

  return (
    <div style={{
      minHeight: "100vh", background: "#0f1117",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'Segoe UI', system-ui, sans-serif"
    }}>
      <div style={{
        background: "linear-gradient(135deg, #1a1d2e, #16213e)",
        border: `1px solid ${error ? "#e74c3c" : "#2a2d3e"}`,
        borderRadius: 20, padding: "44px 40px", width: 340,
        boxShadow: error ? "0 0 30px #e74c3c30" : "0 20px 60px #00000060",
        transition: "border-color 0.3s, box-shadow 0.3s",
        animation: error ? "shake 0.4s" : "none"
      }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{
            width: 72, height: 72, borderRadius: 18, background: "#fff",
            display: "flex", alignItems: "center", justifyContent: "center",
            margin: "0 auto 16px", padding: 6, boxSizing: "border-box",
            boxShadow: "0 4px 20px #0008"
          }}>
            <svg viewBox="0 0 200 200" width="60" height="60" xmlns="http://www.w3.org/2000/svg">
              <path d="M100 15 C52 15 15 52 15 100 C15 148 52 185 100 185 C128 185 153 172 170 151 L150 135 C137 151 120 160 100 160 C67 160 40 133 40 100 C40 67 67 40 100 40 C120 40 137 49 150 65 L170 49 C153 28 128 15 100 15 Z" fill="#009ede"/>
              <text x="108" y="128" textAnchor="middle" fontSize="82" fontWeight="900" fontFamily="Arial Black, sans-serif" fill="#009ede">Z</text>
            </svg>
          </div>
          <div style={{ fontSize: 19, fontWeight: 700, color: "#e8eaf0" }}>Base de Datos de Costeo</div>
          <div style={{ fontSize: 12, color: "#8892b0", marginTop: 6 }}>Ingresa tu clave para continuar</div>
        </div>

        <div style={{ position: "relative", marginBottom: 16 }}>
          <input
            type={show ? "text" : "password"}
            value={input}
            onChange={e => setInput(e.target.value)}
            onKeyDown={e => e.key === "Enter" && intentar()}
            placeholder="Contraseña"
            autoFocus
            style={{
              width: "100%", padding: "12px 44px 12px 16px", borderRadius: 10,
              background: "#0f1117", border: `1px solid ${error ? "#e74c3c" : "#2a2d3e"}`,
              color: "#e8eaf0", fontSize: 15, outline: "none", boxSizing: "border-box",
              transition: "border-color 0.2s"
            }}
          />
          <span
            onClick={() => setShow(!show)}
            style={{
              position: "absolute", right: 14, top: "50%", transform: "translateY(-50%)",
              cursor: "pointer", color: "#8892b0", fontSize: 16, userSelect: "none"
            }}
          >{show ? "🙈" : "👁"}</span>
        </div>

        {error && (
          <div style={{ color: "#e74c3c", fontSize: 12, textAlign: "center", marginBottom: 12 }}>
            Clave incorrecta. Intenta nuevamente.
          </div>
        )}

        <button
          onClick={intentar}
          style={{
            width: "100%", padding: "13px", borderRadius: 10, border: "none",
            background: "linear-gradient(135deg, #3498db, #1abc9c)",
            color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer",
            transition: "opacity 0.2s"
          }}
          onMouseEnter={e => e.currentTarget.style.opacity = "0.88"}
          onMouseLeave={e => e.currentTarget.style.opacity = "1"}
        >
          Ingresar
        </button>
      </div>
      <style>{`@keyframes shake { 0%,100%{transform:translateX(0)} 25%{transform:translateX(-8px)} 75%{transform:translateX(8px)} }`}</style>
    </div>
  );
}

const DATA = {
  "Codos": {
    "2200": { "fecha": "2020-01-08", "cambio": 700, "materiales": [{ "material": "Aceros difusor y piezas", "cantidad": 1, "precio_unit": 1500, "valor_total": 1950 }, { "material": "Arenado", "cantidad": 1, "precio_unit": 65, "valor_total": 572 }, { "material": "Pintura", "cantidad": 1, "precio_unit": 65, "valor_total": 910 }, { "material": "Insumos Soldadura, Corte, Gases", "cantidad": 6, "precio_unit": 1, "valor_total": 420 }, { "material": "Embalaje", "cantidad": 6, "precio_unit": 1, "valor_total": 420 }, { "material": "Tornilleria/otros", "cantidad": 6, "precio_unit": 1, "valor_total": 420 }], "precio_venta": 135.14 },
    "3000": { "fecha": "2020-01-08", "cambio": 750, "materiales": [{ "material": "Aceros difusor y piezas", "cantidad": 1, "precio_unit": 2200, "valor_total": 2860 }, { "material": "Arenado", "cantidad": 1, "precio_unit": 120, "valor_total": 1056 }, { "material": "Pintura", "cantidad": 1, "precio_unit": 120, "valor_total": 1680 }, { "material": "Insumos Soldadura, Corte, Gases", "cantidad": 6, "precio_unit": 1, "valor_total": 420 }, { "material": "Embalaje", "cantidad": 6, "precio_unit": 1, "valor_total": 420 }, { "material": "Tornilleria/otros", "cantidad": 6, "precio_unit": 1, "valor_total": 420 }], "precio_venta": 135.14 },
    "3200": { "fecha": "2020-01-08", "cambio": 700, "materiales": [{ "material": "Aceros difusor y piezas", "cantidad": 1, "precio_unit": 2400, "valor_total": 3120 }, { "material": "Arenado", "cantidad": 1, "precio_unit": 240, "valor_total": 2112 }, { "material": "Pintura", "cantidad": 1, "precio_unit": 240, "valor_total": 3360 }, { "material": "Insumos Soldadura, Corte, Gases", "cantidad": 6, "precio_unit": 1.2, "valor_total": 504 }, { "material": "Embalaje", "cantidad": 6, "precio_unit": 1.2, "valor_total": 504 }, { "material": "Tornilleria/otros", "cantidad": 6, "precio_unit": 1.2, "valor_total": 504 }], "precio_venta": 135.14 },
    "3600": { "fecha": "2020-01-08", "cambio": 750, "materiales": [{ "material": "Aceros difusor y piezas", "cantidad": 1, "precio_unit": 4800, "valor_total": 6240 }, { "material": "Arenado", "cantidad": 1, "precio_unit": 300, "valor_total": 2640 }, { "material": "Pintura", "cantidad": 1, "precio_unit": 300, "valor_total": 4200 }, { "material": "Insumos Soldadura, Corte, Gases", "cantidad": 6, "precio_unit": 1, "valor_total": 420 }, { "material": "Embalaje", "cantidad": 6, "precio_unit": 1, "valor_total": 420 }, { "material": "Tornilleria/otros", "cantidad": 6, "precio_unit": 1, "valor_total": 420 }], "precio_venta": 135.14 },
    "4000": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 1, "precio_unit": 8000, "valor_total": 10400 }, { "material": "Arenado", "cantidad": 1, "precio_unit": 230, "valor_total": 2024 }, { "material": "Pintura", "cantidad": 4, "precio_unit": 57.5, "valor_total": 3220 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 2.8, "precio_unit": 80, "valor_total": 246.4 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 9, "precio_unit": 15, "valor_total": 148.5 }, { "material": "Embalaje", "cantidad": 1, "precio_unit": 1, "valor_total": 500 }], "precio_venta": 34574.04 },
    "4200": { "fecha": null, "cambio": 840, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 1, "precio_unit": 8900, "valor_total": 22250 }, { "material": "Arenado", "cantidad": 1, "precio_unit": 250, "valor_total": 2200 }, { "material": "Pintura", "cantidad": 4, "precio_unit": 69, "valor_total": 3864 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 2.8, "precio_unit": 90, "valor_total": 277.2 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 9, "precio_unit": 15, "valor_total": 148.5 }, { "material": "Embalaje", "cantidad": 1, "precio_unit": 1, "valor_total": 500 }], "precio_venta": 52383.34 }
  },
  "Guillotinas": {
    "G,200": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 2100, "precio_unit": 2.5, "valor_total": 5250 }, { "material": "Arenado", "cantidad": 85, "precio_unit": 8, "valor_total": 680 }, { "material": "Pintura", "cantidad": 85, "precio_unit": 14, "valor_total": 1190 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 85, "precio_unit": 1, "valor_total": 340 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 40, "precio_unit": 1, "valor_total": 320 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 100, "valor_total": 220 }], "precio_venta": 36391.89, "precio_nacional_clp": 19509055.45 },
    "G,220": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 2600, "precio_unit": 2.5, "valor_total": 6500 }, { "material": "Arenado", "cantidad": 100, "precio_unit": 8, "valor_total": 800 }, { "material": "Pintura", "cantidad": 100, "precio_unit": 14, "valor_total": 1400 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 100, "precio_unit": 1, "valor_total": 400 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 50, "precio_unit": 1, "valor_total": 450 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 100, "valor_total": 220 }], "precio_venta": 39022.97, "precio_nacional_clp": 20919531.9 }
  },
  "JZR": {
    "JZR 7": { "fecha": "2020-01-09", "cambio": 700, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 130, "precio_unit": 1.3, "valor_total": 169 }, { "material": "RODETE JZR 7-30/2 250° - 2H", "cantidad": 1, "precio_unit": null, "valor_total": null }, { "material": "Galvanizado", "cantidad": 130, "precio_unit": 0.6, "valor_total": 78 }, { "material": "Pintura", "cantidad": 7, "precio_unit": 14, "valor_total": 98 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 7, "precio_unit": 5, "valor_total": 35 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 70, "valor_total": 70 }, { "material": "Silenciador JAR-7", "cantidad": 2, "precio_unit": 722, "valor_total": 1444 }, { "material": "Rejilla", "cantidad": 2, "precio_unit": 100, "valor_total": 200 }], "precio_venta": 4910, "precio_nacional_clp": 3437000 },
    "JZR 10": { "fecha": "2020-01-09", "cambio": 700, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 200, "precio_unit": 1.3, "valor_total": 260 }, { "material": "RODETE JZR 10 - 250° - 2H", "cantidad": 1, "precio_unit": null, "valor_total": null }, { "material": "Galvanizado", "cantidad": 200, "precio_unit": 0.6, "valor_total": 120 }, { "material": "Pintura", "cantidad": 7, "precio_unit": 14, "valor_total": 98 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 15, "precio_unit": 5, "valor_total": 75 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 70, "valor_total": 70 }, { "material": "Silenciador JAR-10", "cantidad": 1, "precio_unit": 800, "valor_total": 1600 }], "precio_venta": 5244.61, "precio_nacional_clp": 4283101.35 },
    "JZR 12": { "fecha": "2020-01-09", "cambio": 700, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 250, "precio_unit": 1.3, "valor_total": 325 }, { "material": "RODETE JZR 12 - 250° - 2H", "cantidad": 1, "precio_unit": null, "valor_total": null }, { "material": "Galvanizado", "cantidad": 250, "precio_unit": 0.6, "valor_total": 150 }, { "material": "Pintura", "cantidad": 10, "precio_unit": 14, "valor_total": 140 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 10, "precio_unit": 5, "valor_total": 50 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 70, "valor_total": 70 }, { "material": "Silenciador JAR-12", "cantidad": 1, "precio_unit": 1000, "valor_total": 2000 }], "precio_venta": 6005.69, "precio_nacional_clp": 4671096.1 },
    "JZR 14": { "fecha": "2020-01-09", "cambio": 700, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 350, "precio_unit": 1.3, "valor_total": 455 }, { "material": "RODETE JZR 14 - 400° - 2H", "cantidad": 1, "precio_unit": null, "valor_total": null }, { "material": "Galvanizado", "cantidad": 350, "precio_unit": 0.6, "valor_total": 210 }, { "material": "Pintura", "cantidad": 12, "precio_unit": 14, "valor_total": 168 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 12, "precio_unit": 5, "valor_total": 60 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 70, "valor_total": 70 }, { "material": "Silenciador JAR-12", "cantidad": 1, "precio_unit": 1200, "valor_total": 2400 }], "precio_venta": 6939.21, "precio_nacional_clp": 5667020.27 },
    "JZR 16": { "fecha": "2020-01-09", "cambio": 700, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 450, "precio_unit": 1.3, "valor_total": 585 }, { "material": "RODETE JZRi 16 - 400° - 2H", "cantidad": 1, "precio_unit": null, "valor_total": null }, { "material": "Galvanizado", "cantidad": 450, "precio_unit": 0.6, "valor_total": 270 }, { "material": "Pintura", "cantidad": 15, "precio_unit": 14, "valor_total": 210 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 15, "precio_unit": 5, "valor_total": 75 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 70, "valor_total": 70 }, { "material": "Silenciador JAR-16", "cantidad": 1, "precio_unit": 1870, "valor_total": 3740 }, { "material": "Rejilla", "cantidad": 1, "precio_unit": 100, "valor_total": 200 }], "precio_venta": 10958.49, "precio_nacional_clp": 7670945.95 }
  },
  "ZVN": {
    "ZVN 1-6": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 300, "precio_unit": 1.25, "valor_total": 375 }, { "material": "RODETE GEL 6", "cantidad": 1, "precio_unit": 700, "valor_total": 700 }, { "material": "Arenado", "cantidad": 7, "precio_unit": 8.8, "valor_total": 61.6 }, { "material": "Pintura", "cantidad": 7, "precio_unit": 14, "valor_total": 98 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 100, "precio_unit": 1.1, "valor_total": 110 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 15, "precio_unit": 1.1, "valor_total": 49.5 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 80, "valor_total": 80 }, { "material": "Rejillas", "cantidad": 1, "precio_unit": 80, "valor_total": 80 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 45, "valor_total": 45 }], "precio_venta": 5273.44, "precio_nacional_clp": 3955077.7 },
    "ZVN 1-7": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 350, "precio_unit": 1.25, "valor_total": 437.5 }, { "material": "RODETE GEL 7", "cantidad": 1, "precio_unit": 900, "valor_total": 900 }, { "material": "Arenado", "cantidad": 9, "precio_unit": 8.8, "valor_total": 79.2 }, { "material": "Pintura", "cantidad": 9, "precio_unit": 14, "valor_total": 126 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 120, "precio_unit": 1.1, "valor_total": 132 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 20, "precio_unit": 1.1, "valor_total": 66 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 90, "valor_total": 90 }, { "material": "Rejillas", "cantidad": 1, "precio_unit": 80, "valor_total": 80 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 45, "valor_total": 45 }], "precio_venta": 8938.65, "precio_nacional_clp": 6911329.76 },
    "ZVN 1-9": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 400, "precio_unit": 2.5, "valor_total": 1000 }, { "material": "RODETE GEL 9", "cantidad": 1.15, "precio_unit": 1300, "valor_total": 1495 }, { "material": "Arenado", "cantidad": 13, "precio_unit": 8.8, "valor_total": 114.4 }, { "material": "Pintura", "cantidad": 13, "precio_unit": 14, "valor_total": 182 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 150, "precio_unit": 1.1, "valor_total": 165 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 25, "precio_unit": 1.1, "valor_total": 82.5 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 110, "valor_total": 110 }, { "material": "Rejillas", "cantidad": 1, "precio_unit": 101, "valor_total": 101 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 45, "valor_total": 45 }], "precio_venta": 5804.56, "precio_nacional_clp": 5804561.04 },
    "ZVN 1-10": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 460, "precio_unit": 1.3, "valor_total": 598 }, { "material": "Rodete ZVN1-10", "cantidad": 1, "precio_unit": 2300, "valor_total": 2300 }, { "material": "Arenado", "cantidad": 24, "precio_unit": 8.8, "valor_total": 211.2 }, { "material": "Pintura", "cantidad": 4.8, "precio_unit": 14, "valor_total": 134.4 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 150, "precio_unit": 1.1, "valor_total": 165 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 15, "precio_unit": 1.1, "valor_total": 66 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 110, "valor_total": 110 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 45, "valor_total": 45 }, { "material": "Rejillas", "cantidad": 1, "precio_unit": 110, "valor_total": 110 }], "precio_venta": 14383.19, "precio_nacional_clp": 10787391.89 },
    "ZVN 1-12": { "fecha": "2020-01-03", "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 630, "precio_unit": 1.3, "valor_total": 819 }, { "material": "Rodete ZVN1-12", "cantidad": 1, "precio_unit": 2900, "valor_total": 2900 }, { "material": "Arenado", "cantidad": 24, "precio_unit": 8.8, "valor_total": 211.2 }, { "material": "Pintura", "cantidad": 4.8, "precio_unit": 14, "valor_total": 134.4 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 150, "precio_unit": 1.1, "valor_total": 198 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 15, "precio_unit": 1.1, "valor_total": 66 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 110, "valor_total": 132 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 45, "valor_total": 54 }, { "material": "Rejillas", "cantidad": 1, "precio_unit": 110, "valor_total": 110 }], "precio_venta": 16563.59, "precio_nacional_clp": 12422695.95 },
    "ZVN 1-14": { "fecha": "2020-01-03", "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 940, "precio_unit": 1.3, "valor_total": 1222 }, { "material": "Rodete ZVN1-14", "cantidad": 1, "precio_unit": 4000, "valor_total": 4000 }, { "material": "Arenado", "cantidad": 35, "precio_unit": 8.8, "valor_total": 308 }, { "material": "Pintura", "cantidad": 7, "precio_unit": 14, "valor_total": 196 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 150, "precio_unit": 1.1, "valor_total": 231 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 15, "precio_unit": 1.1, "valor_total": 148.5 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 110, "valor_total": 154 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 45, "valor_total": 63 }, { "material": "Rejillas", "cantidad": 1, "precio_unit": 125, "valor_total": 125 }], "precio_venta": 17836.37, "precio_nacional_clp": 13377280.41 },
    "ZVN 1-16": { "fecha": "2020-01-03", "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 1260, "precio_unit": 2.5, "valor_total": 3150 }, { "material": "Rodete ZVN1-16", "cantidad": 1, "precio_unit": 4000, "valor_total": 4000 }, { "material": "Arenado", "cantidad": 43, "precio_unit": 8.8, "valor_total": 378.4 }, { "material": "Pintura", "cantidad": 8.6, "precio_unit": 14, "valor_total": 240.8 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 150, "precio_unit": 1.1, "valor_total": 264 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 15, "precio_unit": 1.1, "valor_total": 148.5 }, { "material": "Rejillas", "cantidad": 1, "precio_unit": 140, "valor_total": 140 }], "precio_venta": 30323.25, "precio_nacional_clp": 22742435.81 },
    "ZVN 1-16 MINA": { "fecha": "2020-01-03", "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 1900, "precio_unit": 1.25, "valor_total": 2375 }, { "material": "Rodete ZVN1-18 (mina)", "cantidad": 1, "precio_unit": 10000, "valor_total": 10000 }, { "material": "Arenado", "cantidad": 43, "precio_unit": 8.8, "valor_total": 378.4 }, { "material": "Pintura", "cantidad": 8.6, "precio_unit": 14, "valor_total": 240.8 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 150, "precio_unit": 1.1, "valor_total": 264 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 15, "precio_unit": 1.1, "valor_total": 148.5 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 110, "valor_total": 176 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 45, "valor_total": 72 }, { "material": "Rejillas", "cantidad": 1, "precio_unit": 140, "valor_total": 140 }], "precio_venta": 52216.45, "precio_nacional_clp": 39162334.46 },
    "ZVN 1-18": { "fecha": "2020-01-03", "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 2000, "precio_unit": 1.3, "valor_total": 2600 }, { "material": "Rodete ZVN1-18", "cantidad": 1, "precio_unit": 5992.59, "valor_total": 5992.59 }, { "material": "Arenado", "cantidad": 80, "precio_unit": 8.8, "valor_total": 704 }, { "material": "Pintura", "cantidad": 20, "precio_unit": 14, "valor_total": 560 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 150, "precio_unit": 1.1, "valor_total": 297 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 15, "precio_unit": 1.1, "valor_total": 148.5 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 110, "valor_total": 198 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 45, "valor_total": 81 }, { "material": "Rejillas", "cantidad": 1, "precio_unit": 140, "valor_total": 140 }], "precio_venta": 44045.87, "precio_nacional_clp": 33034401.84 },
    "ZVN 1-20": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 3000, "precio_unit": 1.3, "valor_total": 3900 }, { "material": "Rodete ZVN1-22", "cantidad": 1, "precio_unit": 15000, "valor_total": 15000 }, { "material": "Arenado", "cantidad": 90, "precio_unit": 8, "valor_total": 720 }, { "material": "Pintura", "cantidad": 90, "precio_unit": 14, "valor_total": 1260 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 200, "precio_unit": 1, "valor_total": 440 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 40, "precio_unit": 1, "valor_total": 360 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 100, "valor_total": 220 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 40, "valor_total": 88 }, { "material": "Directrices", "cantidad": 1, "precio_unit": 880, "valor_total": 880 }], "precio_venta": 59511.89, "precio_nacional_clp": 36407275.04 },
    "ZVN 1-22": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 3400, "precio_unit": 1.3, "valor_total": 4420 }, { "material": "Rodete ZVN1-22", "cantidad": 1, "precio_unit": 15000, "valor_total": 15000 }, { "material": "Arenado", "cantidad": 94, "precio_unit": 8, "valor_total": 752 }, { "material": "Pintura", "cantidad": 100, "precio_unit": 14, "valor_total": 1400 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 200, "precio_unit": 1, "valor_total": 440 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 40, "precio_unit": 1, "valor_total": 360 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 100, "valor_total": 220 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 40, "valor_total": 88 }, { "material": "Directrices", "cantidad": 1, "precio_unit": 880, "valor_total": 880 }], "precio_venta": 115310.81, "precio_nacional_clp": 61816104.76 },
    "ZVN 1-24": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 5000, "precio_unit": 1.3, "valor_total": 6500 }, { "material": "Rodete ZVN1-24", "cantidad": 1, "precio_unit": 20000, "valor_total": 20000 }, { "material": "Arenado", "cantidad": 150, "precio_unit": 8, "valor_total": 1200 }, { "material": "Pintura", "cantidad": 150, "precio_unit": 14, "valor_total": 2100 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 200, "precio_unit": 1, "valor_total": 480 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 40, "precio_unit": 1, "valor_total": 360 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 100, "valor_total": 240 }, { "material": "Tensor apoyo", "cantidad": 1, "precio_unit": 40, "valor_total": 96 }, { "material": "Directrices", "cantidad": 1, "precio_unit": 880, "valor_total": 880 }], "precio_venta": 103088.65, "precio_nacional_clp": 63065996.82 },
    "ZVN 1-28": { "fecha": "2020-01-15", "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 9000, "precio_unit": 1.3, "valor_total": 11700 }, { "material": "Rodete s/oferta Jose Angel", "cantidad": 1, "precio_unit": 13961.22, "valor_total": 13961.22 }, { "material": "Arenado", "cantidad": 200, "precio_unit": 8.8, "valor_total": 1760 }, { "material": "Pintura", "cantidad": 50, "precio_unit": 14, "valor_total": 1400 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 150, "precio_unit": 1.1, "valor_total": 462 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 15, "precio_unit": 1.1, "valor_total": 148.5 }, { "material": "Embalaje", "cantidad": 1, "precio_unit": 500, "valor_total": 500 }], "precio_venta": 116592.93, "precio_nacional_clp": 87444698.15 },
    "ZVN 1-30": { "fecha": "2020-01-15", "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 11000, "precio_unit": 1.3, "valor_total": 14300 }, { "material": "Rodete ZVN 1-30 s/ANDINA", "cantidad": 1, "precio_unit": 35555.56, "valor_total": 35555.56 }, { "material": "Arenado", "cantidad": 280, "precio_unit": 8.8, "valor_total": 2464 }, { "material": "Pintura", "cantidad": 70, "precio_unit": 14, "valor_total": 1960 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 150, "precio_unit": 1.1, "valor_total": 462 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 15, "precio_unit": 1.1, "valor_total": 148.5 }, { "material": "Embalaje", "cantidad": 1, "precio_unit": 500, "valor_total": 500 }], "precio_venta": 206296.63, "precio_nacional_clp": 154722471.85 }
  },
  "Piezas Especiales": {
    "GV 160": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 1500, "precio_unit": 2.5, "valor_total": 3750 }, { "material": "Arenado", "cantidad": 75, "precio_unit": 8, "valor_total": 600 }, { "material": "Pintura", "cantidad": 75, "precio_unit": 14, "valor_total": 1050 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 75, "precio_unit": 1, "valor_total": 300 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 40, "precio_unit": 1, "valor_total": 320 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 100, "valor_total": 220 }], "precio_venta": 27654.05, "precio_nacional_clp": 14824853.72 },
    "GH 200": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 2100, "precio_unit": 2.5, "valor_total": 5250 }, { "material": "Arenado", "cantidad": 85, "precio_unit": 8, "valor_total": 680 }, { "material": "Pintura", "cantidad": 85, "precio_unit": 14, "valor_total": 1190 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 85, "precio_unit": 1, "valor_total": 340 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 40, "precio_unit": 1, "valor_total": 320 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 100, "valor_total": 220 }], "precio_venta": 36391.89, "precio_nacional_clp": 19509055.45 },
    "GH 220": { "fecha": null, "cambio": 750, "materiales": [{ "material": "Aceros ventiladores y piezas", "cantidad": 2600, "precio_unit": 2.5, "valor_total": 6500 }, { "material": "Arenado", "cantidad": 100, "precio_unit": 8, "valor_total": 800 }, { "material": "Pintura", "cantidad": 100, "precio_unit": 14, "valor_total": 1400 }, { "material": "Insumos Soldadura, Corte, Gases, tornilleria", "cantidad": 100, "precio_unit": 1, "valor_total": 400 }, { "material": "Otros Cortes y Piezas Exterior", "cantidad": 50, "precio_unit": 1, "valor_total": 450 }, { "material": "Otros Menores Embalaje", "cantidad": 1, "precio_unit": 100, "valor_total": 220 }], "precio_venta": 39022.97, "precio_nacional_clp": 20919531.9 },
    "DIF 220-290-300 + BULBO": { "fecha": "2020-01-08", "cambio": 750, "materiales": [{ "material": "Aceros difusor y piezas", "cantidad": 2800, "precio_unit": 2.5, "valor_total": 7000 }, { "material": "CHAPA PERFORADA 1500x2000", "cantidad": 16, "precio_unit": 100, "valor_total": 1600 }, { "material": "LANA DE ROCA", "cantidad": 70, "precio_unit": 9, "valor_total": 630 }, { "material": "Arenado", "cantidad": 240, "precio_unit": 8.8, "valor_total": 2112 }, { "material": "Pintura", "cantidad": 240, "precio_unit": 14, "valor_total": 3360 }, { "material": "Insumos Soldadura, Corte, Gases", "cantidad": 1, "precio_unit": 70, "valor_total": 420 }, { "material": "Embalaje", "cantidad": 1, "precio_unit": 70, "valor_total": 420 }, { "material": "Tornilleria/otros", "cantidad": 1, "precio_unit": 70, "valor_total": 420 }], "precio_venta": null }
  },
  "Silenciadores": {
    "FAR 60/100": { "fecha": null, "cambio": 750, "materiales": [], "precio_venta": 444.41 },
    "FAR 70/100": { "fecha": null, "cambio": 680, "materiales": [], "precio_venta": 510.32 },
    "FAR 90/100": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 624.31 },
    "FAR 100/100": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 720.1 },
    "FAR 120/200": { "fecha": null, "cambio": 750, "materiales": [], "precio_venta": 936.75 },
    "FAR 140/200": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 1529.83 },
    "FAR 160/200": { "fecha": null, "cambio": 680, "materiales": [], "precio_venta": 1740.47 },
    "FAR 180/200": { "fecha": null, "cambio": 520, "materiales": [], "precio_venta": 1696.88 },
    "FAR 200/200": { "fecha": null, "cambio": 750, "materiales": [], "precio_venta": 1891.17 },
    "FAR 230/200": { "fecha": null, "cambio": 750, "materiales": [], "precio_venta": 2213.26 },
    "FAR 250/250": { "fecha": null, "cambio": 520, "materiales": [], "precio_venta": 2934.61 },
    "VAV 18": { "fecha": null, "cambio": 750, "materiales": [], "precio_venta": 417.19 },
    "VAV 21": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 558.38 },
    "VAV 25": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 577.99 },
    "VAV 27": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 732.16 },
    "VAV 32": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 839.91 },
    "VAV 33": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 944.89 },
    "VAV 35": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 954.89 },
    "VAV 37": { "fecha": null, "cambio": 700, "materiales": [], "precio_venta": 1038.01 }
  },
  "Difusores": {
    "TOBERA AERODINAMICA": { "fecha": "2019-01-03", "cambio": 680, "modelos": [{ "diametro": "DIAM. 600MM", "costo_usd": 329.33, "venta_usd_nacional": 445.04, "venta_clp_nacional": 302625.16 }, { "diametro": "DIAM. 700MM", "costo_usd": 340.95, "venta_usd_nacional": 460.74, "venta_clp_nacional": 313306.31 }, { "diametro": "DIAM. 900MM", "costo_usd": 379.73, "venta_usd_nacional": 513.15, "venta_clp_nacional": 348941.26 }, { "diametro": "DIAM. 1000MM", "costo_usd": 407.85, "venta_usd_nacional": 551.15, "venta_clp_nacional": 374784.58 }, { "diametro": "DIAM. 1200MM", "costo_usd": 496.96, "venta_usd_nacional": 671.57, "venta_clp_nacional": 456670.12 }, { "diametro": "DIAM. 1400MM", "costo_usd": 604.25, "venta_usd_nacional": 816.56, "venta_clp_nacional": 555258.37 }, { "diametro": "DIAM. 1600MM", "costo_usd": 432.58, "venta_usd_nacional": 584.57, "venta_clp_nacional": 397505 }] },
    "TOBERA CONICA": { "fecha": "2019-01-03", "cambio": 800, "modelos": [{ "diametro": "DIAM. 900MM", "costo_usd": 910.9, "venta_usd_nacional": 1230.94, "venta_clp_nacional": 984755.2 }, { "diametro": "DIAM. 1000MM", "costo_usd": 921.22, "venta_usd_nacional": 1244.89, "venta_clp_nacional": 995915.69 }, { "diametro": "DIAM. 1200MM", "costo_usd": 962.66, "venta_usd_nacional": 1300.89, "venta_clp_nacional": 1040712.35 }, { "diametro": "DIAM. 1400MM", "costo_usd": 1004.1, "venta_usd_nacional": 1356.89, "venta_clp_nacional": 1085509.02 }, { "diametro": "DIAM. 1600MM", "costo_usd": 1071.82, "venta_usd_nacional": 1448.41, "venta_clp_nacional": 1158727.3 }, { "diametro": "DIAM. 1800MM", "costo_usd": 1156.05, "venta_usd_nacional": 1562.23, "venta_clp_nacional": 1249783.42 }, { "diametro": "DIAM. 2000MM", "costo_usd": 2462.3, "venta_usd_nacional": 3327.43, "venta_clp_nacional": 2661943.49 }, { "diametro": "DIAM. 2200MM", "costo_usd": 2507.69, "venta_usd_nacional": 3388.78, "venta_clp_nacional": 2711021.23 }, { "diametro": "DIAM. 2500MM", "costo_usd": 2607.75, "venta_usd_nacional": 3523.98, "venta_clp_nacional": 2819183.79 }, { "diametro": "DIAM. 3000MM", "costo_usd": 3871.36, "venta_usd_nacional": 5231.56, "venta_clp_nacional": 4185250.7 }, { "diametro": "DIAM. 3300MM", "costo_usd": 4140.2, "venta_usd_nacional": 5594.86, "venta_clp_nacional": 4475888.54 }] },
    "DIFUSORES": { "fecha": "2019-01-03", "cambio": 800, "modelos": [{ "diametro": "DIAM. 900MM", "costo_usd": 2760.9, "venta_usd_nacional": 4313.9, "venta_clp_nacional": 3451123.15 }, { "diametro": "DIAM. 1000MM", "costo_usd": 3198.76, "venta_usd_nacional": 4998.06, "venta_clp_nacional": 3998444.01 }, { "diametro": "DIAM. 1200MM", "costo_usd": 3501.72, "venta_usd_nacional": 5471.44, "venta_clp_nacional": 4377151.25 }, { "diametro": "DIAM. 1400MM", "costo_usd": 4477.9, "venta_usd_nacional": 6996.72, "venta_clp_nacional": 5597376.31 }, { "diametro": "DIAM. 1600MM", "costo_usd": 4703.44, "venta_usd_nacional": 7349.12, "venta_clp_nacional": 5879296.46 }, { "diametro": "DIAM. 1800MM", "costo_usd": 5449.09, "venta_usd_nacional": 8514.21, "venta_clp_nacional": 6811364.79 }, { "diametro": "DIAM. 2000MM", "costo_usd": 9160.72, "venta_usd_nacional": 14313.63, "venta_clp_nacional": 11450901.68 }, { "diametro": "DIAM. 2200MM", "costo_usd": 9401.51, "venta_usd_nacional": 14689.87, "venta_clp_nacional": 11751893.41 }, { "diametro": "DIAM. 2500MM", "costo_usd": 5883.05, "venta_usd_nacional": 9192.27, "venta_clp_nacional": 7353812.26 }] }
  },
  "Adaptación a Manga": {
    "TIPO A": { "fecha": "2019-01-03", "cambio": 750, "modelos": [{ "diametro": "DIAM. 500MM", "costo_usd": 181.61, "venta_usd_nacional": 283.77, "venta_clp_nacional": 212828.69 }, { "diametro": "DIAM. 600MM", "costo_usd": 188.52, "venta_usd_nacional": 294.56, "venta_clp_nacional": 220916.3 }, { "diametro": "DIAM. 700MM", "costo_usd": 195.42, "venta_usd_nacional": 305.34, "venta_clp_nacional": 229003.91 }, { "diametro": "DIAM. 900MM", "costo_usd": 209.22, "venta_usd_nacional": 326.91, "venta_clp_nacional": 245179.14 }, { "diametro": "DIAM. 1000MM", "costo_usd": 260.12, "venta_usd_nacional": 406.44, "venta_clp_nacional": 304829.25 }, { "diametro": "DIAM. 1200MM", "costo_usd": 273.92, "venta_usd_nacional": 428.01, "venta_clp_nacional": 321004.47 }, { "diametro": "DIAM. 1400MM", "costo_usd": 331.73, "venta_usd_nacional": 518.32, "venta_clp_nacional": 388742.2 }, { "diametro": "DIAM. 1600MM", "costo_usd": 345.53, "venta_usd_nacional": 539.89, "venta_clp_nacional": 404917.42 }, { "diametro": "DIAM. 1800MM", "costo_usd": 359.33, "venta_usd_nacional": 561.46, "venta_clp_nacional": 421092.65 }, { "diametro": "DIAM. 2000MM", "costo_usd": 668.27, "venta_usd_nacional": 1044.17, "venta_clp_nacional": 783125.9 }, { "diametro": "DIAM. 2200MM", "costo_usd": 757.67, "venta_usd_nacional": 1183.86, "venta_clp_nacional": 887894.75 }, { "diametro": "DIAM. 2500MM", "costo_usd": 803.77, "venta_usd_nacional": 1255.9, "venta_clp_nacional": 941923.02 }, { "diametro": "DIAM. 3200MM", "costo_usd": 911.35, "venta_usd_nacional": 1423.99, "venta_clp_nacional": 1067988.98 }] },
    "TIPO B": { "fecha": "2019-01-03", "cambio": 700, "modelos": [{ "diametro": "DIAM. 600MM", "costo_usd": 508.44, "venta_usd_nacional": 687.08, "venta_clp_nacional": 480952.73 }, { "diametro": "DIAM. 700MM", "costo_usd": 506.43, "venta_usd_nacional": 684.36, "venta_clp_nacional": 479052.33 }, { "diametro": "DIAM. 900MM", "costo_usd": 535.59, "venta_usd_nacional": 723.77, "venta_clp_nacional": 506638.96 }, { "diametro": "DIAM. 1000MM", "costo_usd": 789.67, "venta_usd_nacional": 1067.12, "venta_clp_nacional": 746985.45 }, { "diametro": "DIAM. 1200MM", "costo_usd": 902.32, "venta_usd_nacional": 1219.36, "venta_clp_nacional": 853549.11 }, { "diametro": "DIAM. 1400MM", "costo_usd": 1225.34, "venta_usd_nacional": 1655.87, "venta_clp_nacional": 1159110.07 }, { "diametro": "DIAM. 1600MM", "costo_usd": 1277.94, "venta_usd_nacional": 1726.94, "venta_clp_nacional": 1208860.21 }, { "diametro": "DIAM. 1800MM", "costo_usd": 1335.15, "venta_usd_nacional": 1804.26, "venta_clp_nacional": 1262980.62 }] },
    "TIPO C": { "fecha": "2019-01-03", "cambio": 700, "modelos": [{ "diametro": "DIAM. 600MM", "costo_usd": 451.97, "venta_usd_nacional": 610.78, "venta_clp_nacional": 427542.72 }, { "diametro": "DIAM. 700MM", "costo_usd": 467.22, "venta_usd_nacional": 631.37, "venta_clp_nacional": 441960.36 }, { "diametro": "DIAM. 900MM", "costo_usd": 533.91, "venta_usd_nacional": 721.5, "venta_clp_nacional": 505047.87 }, { "diametro": "DIAM. 1000MM", "costo_usd": 715.54, "venta_usd_nacional": 966.94, "venta_clp_nacional": 676859.63 }, { "diametro": "DIAM. 1200MM", "costo_usd": 808.39, "venta_usd_nacional": 1092.42, "venta_clp_nacional": 764693.56 }, { "diametro": "DIAM. 1400MM", "costo_usd": 1079.6, "venta_usd_nacional": 1458.92, "venta_clp_nacional": 1021245.05 }, { "diametro": "DIAM. 1600MM", "costo_usd": 1200.17, "venta_usd_nacional": 1621.86, "venta_clp_nacional": 1135300.6 }, { "diametro": "DIAM. 1800MM", "costo_usd": 1258.71, "venta_usd_nacional": 1700.96, "venta_clp_nacional": 1190669.66 }] },
    "PANTALON": { "fecha": "2019-01-03", "cambio": 840, "modelos": [{ "diametro": "DIAM. 700MM", "costo_usd": 1885.63, "venta_usd_nacional": 2548.15, "venta_clp_nacional": 2140442.87 }, { "diametro": "DIAM. 900MM", "costo_usd": 2084.83, "venta_usd_nacional": 2817.33, "venta_clp_nacional": 2366561.29 }, { "diametro": "DIAM. 1000MM", "costo_usd": 2096.36, "venta_usd_nacional": 2832.92, "venta_clp_nacional": 2379651.12 }, { "diametro": "DIAM. 1200MM", "costo_usd": 2272.96, "venta_usd_nacional": 3071.56, "venta_clp_nacional": 2580112.66 }, { "diametro": "DIAM. 1400MM", "costo_usd": 2324.44, "venta_usd_nacional": 3141.14, "venta_clp_nacional": 2638555.27 }, { "diametro": "DIAM. 4200MM", "costo_usd": 10309, "venta_usd_nacional": 13931.08, "venta_clp_nacional": 11702109.83 }] }
  },
  "Válvulas Manga": {
    "Tipo T": { "fecha": "2019-01-03", "cambio": 750, "modelos": [{ "diametro": "T", "costo_usd": 1312.1, "venta_usd_nacional": 1773.11, "venta_clp_nacional": 1329829.47 }] },
    "Tipo Y": { "fecha": "2019-01-03", "cambio": 750, "modelos": [{ "diametro": "Y", "costo_usd": 1378.81, "venta_usd_nacional": 1863.25, "venta_clp_nacional": 1397440.74 }] }
  }
};

const fmt = (v, decimals = 0) => {
  if (v == null || isNaN(v)) return "—";
  return new Intl.NumberFormat("es-CL", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(v);
};

const ICONS = {
  "Codos": "↩", "Guillotinas": "⊟", "JZR": "⟳", "ZVN": "⬡",
  "Piezas Especiales": "◈", "Silenciadores": "⊕", "Difusores": "◎",
  "Adaptación a Manga": "⊂", "Válvulas Manga": "⊡"
};

const COLORS = [
  "#e74c3c","#e67e22","#f1c40f","#2ecc71","#1abc9c",
  "#3498db","#9b59b6","#e91e63","#00bcd4"
];

export default function App() {
  const [autenticado, setAutenticado] = useState(false);
  const [cat, setCat] = useState(null);
  const [modelo, setModelo] = useState(null);

  if (!autenticado) return <Login onAcceso={() => setAutenticado(true)} />;

  const categorias = Object.keys(DATA);
  const catColor = cat ? COLORS[categorias.indexOf(cat) % COLORS.length] : "#3498db";

  const modelos = cat ? Object.keys(DATA[cat]) : [];
  const datos = cat && modelo ? DATA[cat][modelo] : null;

  const hasMateriales = datos && datos.materiales && datos.materiales.length > 0;
  const hasModelos = datos && datos.modelos && datos.modelos.length > 0;

  const totalMat = hasMateriales
    ? datos.materiales.reduce((s, m) => s + (m.valor_total || 0), 0)
    : 0;

  return (
    <div style={{
      minHeight: "100vh", background: "#0f1117", color: "#e8eaf0",
      fontFamily: "'Segoe UI', system-ui, sans-serif", display: "flex", flexDirection: "column"
    }}>
      {/* HEADER */}
      <div style={{
        background: "linear-gradient(135deg, #1a1d2e 0%, #16213e 100%)",
        borderBottom: "1px solid #2a2d3e", padding: "18px 28px",
        display: "flex", alignItems: "center", gap: 14
      }}>
        <div style={{
          width: 40, height: 40, borderRadius: 10, background: "#fff",
          display: "flex", alignItems: "center", justifyContent: "center",
          flexShrink: 0, padding: 3, boxSizing: "border-box",
          boxShadow: "0 2px 10px #0006"
        }}>
          <svg viewBox="0 0 200 200" width="34" height="34" xmlns="http://www.w3.org/2000/svg">
            <path d="M100 15 C52 15 15 52 15 100 C15 148 52 185 100 185 C128 185 153 172 170 151 L150 135 C137 151 120 160 100 160 C67 160 40 133 40 100 C40 67 67 40 100 40 C120 40 137 49 150 65 L170 49 C153 28 128 15 100 15 Z" fill="#009ede"/>
            <text x="108" y="128" textAnchor="middle" fontSize="82" fontWeight="900" fontFamily="Arial Black, sans-serif" fill="#009ede">Z</text>
          </svg>
        </div>
        <div>
          <div style={{ fontSize: 17, fontWeight: 700, letterSpacing: 0.3 }}>Base de Datos de Costeo</div>
          <div style={{ fontSize: 11, color: "#8892b0", marginTop: 1 }}>
            {categorias.length} categorías · {categorias.reduce((s, c) => s + Object.keys(DATA[c]).length, 0)} modelos
          </div>
        </div>
        {cat && (
          <button onClick={() => { setCat(null); setModelo(null); }} style={{
            marginLeft: "auto", background: "rgba(255,255,255,0.07)", border: "1px solid #2a2d3e",
            color: "#8892b0", borderRadius: 8, padding: "5px 14px", cursor: "pointer", fontSize: 12
          }}>← Inicio</button>
        )}
      </div>

      <div style={{ flex: 1, padding: "24px 28px", maxWidth: 1100, margin: "0 auto", width: "100%" }}>

        {/* VISTA INICIO: CATEGORÍAS */}
        {!cat && (
          <div>
            <div style={{ fontSize: 13, color: "#8892b0", marginBottom: 16 }}>Selecciona una categoría de producto</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 14 }}>
              {categorias.map((c, i) => (
                <button key={c} onClick={() => { setCat(c); setModelo(null); }} style={{
                  background: "linear-gradient(135deg, #1a1d2e, #16213e)",
                  border: `1px solid ${COLORS[i % COLORS.length]}40`,
                  borderRadius: 14, padding: "22px 20px", cursor: "pointer",
                  color: "#e8eaf0", textAlign: "left", transition: "all 0.18s",
                  position: "relative", overflow: "hidden"
                }}
                  onMouseEnter={e => { e.currentTarget.style.borderColor = COLORS[i % COLORS.length]; e.currentTarget.style.transform = "translateY(-2px)"; }}
                  onMouseLeave={e => { e.currentTarget.style.borderColor = `${COLORS[i % COLORS.length]}40`; e.currentTarget.style.transform = "none"; }}
                >
                  <div style={{ fontSize: 28, marginBottom: 10 }}>{ICONS[c] || "◆"}</div>
                  <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>{c}</div>
                  <div style={{ fontSize: 11, color: "#8892b0" }}>{Object.keys(DATA[c]).length} modelos</div>
                  <div style={{
                    position: "absolute", bottom: 0, right: 0, width: 60, height: 60,
                    background: `${COLORS[i % COLORS.length]}12`, borderRadius: "50% 0 14px 0"
                  }} />
                </button>
              ))}
            </div>
          </div>
        )}

        {/* VISTA CATEGORÍA: MODELOS */}
        {cat && !modelo && (
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
              <span style={{ fontSize: 22 }}>{ICONS[cat]}</span>
              <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700 }}>{cat}</h2>
              <span style={{
                marginLeft: 6, background: `${catColor}22`, color: catColor,
                borderRadius: 20, padding: "2px 12px", fontSize: 11, fontWeight: 600
              }}>{modelos.length} modelos</span>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))", gap: 10 }}>
              {modelos.map(m => {
                const d = DATA[cat][m];
                const pv = d.precio_venta;
                const hasModelos2 = d.modelos && d.modelos.length > 0;
                return (
                  <button key={m} onClick={() => setModelo(m)} style={{
                    background: "#1a1d2e", border: `1px solid #2a2d3e`,
                    borderRadius: 12, padding: "16px 14px", cursor: "pointer",
                    color: "#e8eaf0", textAlign: "left", transition: "all 0.15s"
                  }}
                    onMouseEnter={e => { e.currentTarget.style.borderColor = catColor; e.currentTarget.style.background = "#1f2235"; }}
                    onMouseLeave={e => { e.currentTarget.style.borderColor = "#2a2d3e"; e.currentTarget.style.background = "#1a1d2e"; }}
                  >
                    <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 6, color: catColor }}>{m}</div>
                    {d.fecha && <div style={{ fontSize: 10, color: "#8892b0", marginBottom: 4 }}>📅 {d.fecha}</div>}
                    {pv && <div style={{ fontSize: 11, color: "#2ecc71", fontWeight: 600 }}>
                      USD {fmt(pv, 2)}
                    </div>}
                    {hasModelos2 && <div style={{ fontSize: 10, color: "#8892b0" }}>{d.modelos.length} diámetros</div>}
                  </button>
                );
              })}
            </div>
          </div>
        )}

        {/* VISTA DETALLE DE MODELO */}
        {cat && modelo && datos && (
          <div>
            {/* Breadcrumb */}
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 18, fontSize: 12, color: "#8892b0" }}>
              <span style={{ cursor: "pointer", color: catColor }} onClick={() => setModelo(null)}>{cat}</span>
              <span>›</span>
              <span style={{ color: "#e8eaf0", fontWeight: 600 }}>{modelo}</span>
            </div>

            {/* Card info general */}
            <div style={{
              background: "linear-gradient(135deg, #1a1d2e, #1f2235)",
              border: `1px solid ${catColor}50`, borderRadius: 16, padding: "20px 24px",
              marginBottom: 20, display: "flex", flexWrap: "wrap", gap: 28, alignItems: "center"
            }}>
              <div>
                <div style={{ fontSize: 11, color: "#8892b0", marginBottom: 4 }}>MODELO</div>
                <div style={{ fontSize: 22, fontWeight: 800, color: catColor }}>{modelo}</div>
              </div>
              {datos.fecha && (
                <div>
                  <div style={{ fontSize: 11, color: "#8892b0", marginBottom: 4 }}>FECHA COSTEO</div>
                  <div style={{ fontSize: 15, fontWeight: 600 }}>{datos.fecha}</div>
                </div>
              )}
              {datos.cambio && (
                <div>
                  <div style={{ fontSize: 11, color: "#8892b0", marginBottom: 4 }}>CAMBIO CLP/USD</div>
                  <div style={{ fontSize: 15, fontWeight: 600 }}>${fmt(datos.cambio)}</div>
                </div>
              )}
              {datos.precio_venta && (
                <div style={{ marginLeft: "auto" }}>
                  <div style={{ fontSize: 11, color: "#8892b0", marginBottom: 4 }}>PRECIO VENTA</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: "#2ecc71" }}>USD {fmt(datos.precio_venta, 2)}</div>
                  {datos.precio_nacional_clp && (
                    <div style={{ fontSize: 11, color: "#8892b0", marginTop: 2 }}>
                      CLP {fmt(datos.precio_nacional_clp)}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* TABLA MATERIALES */}
            {hasMateriales && (
              <div>
                <div style={{ fontSize: 13, fontWeight: 700, color: "#8892b0", marginBottom: 12, textTransform: "uppercase", letterSpacing: 1 }}>
                  Materiales
                </div>
                <div style={{ background: "#1a1d2e", borderRadius: 14, overflow: "hidden", border: "1px solid #2a2d3e" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: "#16192a" }}>
                        {["Material", "Cantidad", "Precio Unit. USD", "Valor Total USD"].map(h => (
                          <th key={h} style={{ padding: "11px 16px", textAlign: h === "Material" ? "left" : "right", fontSize: 11, color: "#8892b0", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.8 }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {datos.materiales.map((m, i) => (
                        <tr key={i} style={{ borderTop: "1px solid #2a2d3e" }}
                          onMouseEnter={e => e.currentTarget.style.background = "#1f2235"}
                          onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                        >
                          <td style={{ padding: "11px 16px", fontSize: 13, fontWeight: 500 }}>{m.material}</td>
                          <td style={{ padding: "11px 16px", fontSize: 13, textAlign: "right", color: "#8892b0" }}>{fmt(m.cantidad, 2)}</td>
                          <td style={{ padding: "11px 16px", fontSize: 13, textAlign: "right", color: "#8892b0" }}>{m.precio_unit ? `${fmt(m.precio_unit, 2)}` : "—"}</td>
                          <td style={{ padding: "11px 16px", fontSize: 13, textAlign: "right", fontWeight: 600, color: m.valor_total ? "#e8eaf0" : "#555" }}>
                            {m.valor_total ? `${fmt(m.valor_total, 2)}` : "—"}
                          </td>
                        </tr>
                      ))}
                      <tr style={{ borderTop: "2px solid #2a2d3e", background: "#16192a" }}>
                        <td colSpan={3} style={{ padding: "11px 16px", fontSize: 12, fontWeight: 700, color: "#8892b0", textTransform: "uppercase" }}>Total Materiales</td>
                        <td style={{ padding: "11px 16px", textAlign: "right", fontWeight: 800, fontSize: 15, color: catColor }}>{fmt(totalMat, 2)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TABLA POR DIÁMETROS */}
            {hasModelos && (
              <div>
                <div style={{ fontSize: 13, fontWeight: 700, color: "#8892b0", marginBottom: 12, textTransform: "uppercase", letterSpacing: 1 }}>
                  Precios por Diámetro
                </div>
                <div style={{ background: "#1a1d2e", borderRadius: 14, overflow: "hidden", border: "1px solid #2a2d3e" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ background: "#16192a" }}>
                        {["Diámetro", "Costo USD", "Venta USD Nacional", "Venta CLP Nacional"].map(h => (
                          <th key={h} style={{ padding: "11px 16px", textAlign: h === "Diámetro" ? "left" : "right", fontSize: 11, color: "#8892b0", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.8 }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {datos.modelos.filter(m => m.costo_usd || m.venta_usd_nacional).map((m, i) => (
                        <tr key={i} style={{ borderTop: "1px solid #2a2d3e" }}
                          onMouseEnter={e => e.currentTarget.style.background = "#1f2235"}
                          onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                        >
                          <td style={{ padding: "11px 16px", fontSize: 13, fontWeight: 600, color: catColor }}>{m.diametro}</td>
                          <td style={{ padding: "11px 16px", fontSize: 13, textAlign: "right", color: "#8892b0" }}>{fmt(m.costo_usd, 2)}</td>
                          <td style={{ padding: "11px 16px", fontSize: 13, textAlign: "right", fontWeight: 600, color: "#2ecc71" }}>{fmt(m.venta_usd_nacional, 2)}</td>
                          <td style={{ padding: "11px 16px", fontSize: 13, textAlign: "right", color: "#8892b0" }}>{fmt(m.venta_clp_nacional)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Solo precio venta (sin materiales ni diámetros) */}
            {!hasMateriales && !hasModelos && datos.precio_venta && (
              <div style={{ background: "#1a1d2e", borderRadius: 14, padding: 24, border: "1px solid #2a2d3e", textAlign: "center" }}>
                <div style={{ fontSize: 12, color: "#8892b0", marginBottom: 8 }}>PRECIO DE VENTA</div>
                <div style={{ fontSize: 32, fontWeight: 800, color: "#2ecc71" }}>USD {fmt(datos.precio_venta, 2)}</div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
