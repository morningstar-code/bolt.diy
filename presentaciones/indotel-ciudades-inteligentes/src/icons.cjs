const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const lu = require("react-icons/lu");
const fa = require("react-icons/fa6");
const fs = require("fs");
const path = require("path");

const OUT = path.join(__dirname, "icons");
fs.mkdirSync(OUT, { recursive: true });

const W = "FFFFFF";
const NAVY = "0D1B5E";
const BLUE = "0B4489";
const GREEN = "2E9E5B";

// El original usa iconografía de trazo fino (Lucide), no siluetas macizas.
// [componente, color, grosorDeTrazo]
const ICONS = {
  // 1 · franja de pilares de la portada
  s1_datos:          [lu.LuDatabase,     W, 1.7],
  s1_interop:        [lu.LuNetwork,      W, 1.7],
  s1_estandares:     [lu.LuShieldCheck,  W, 1.7],
  s1_conectividad:   [lu.LuWifi,         W, 1.7],
  s1_sostenibilidad: [lu.LuLeaf,         W, 1.7],

  // 2 · tarjeta flotante
  s2_servicios:      [lu.LuUserCog,      W, 1.9],
  s2_decisiones:     [lu.LuWorkflow,     W, 1.9],
  s2_inversion:      [lu.LuLandmark,     W, 1.9],
  s2_ciudades:       [lu.LuBuilding2,    W, 1.9],

  // 2 · banda de indicadores
  s2_kpi_municipios: [lu.LuNetwork,        W, 1.6],
  s2_kpi_ciudadanos: [lu.LuUsers,          W, 1.6],
  s2_kpi_proyectos:  [lu.LuClipboardCheck, W, 1.6],
  s2_kpi_decision:   [lu.LuLightbulb,      W, 1.6],

  // 5 · tres columnas
  s5_naturaleza:     [lu.LuLandmark,    W, 1.7],
  s5_tecnico:        [lu.LuShuffle,     W, 1.9],
  s5_valor:          [lu.LuCheck,       W, 2.4],

  // 6 · cómo opera OASC
  s6_gobernanza:     [lu.LuLandmark,    W, 1.9],
  s6_servicios:      [lu.LuBookOpen,    W, 1.9],
  s6_proyectos:      [lu.LuWorkflow,    W, 1.9],
  s6_itu:            [lu.LuHandshake,   W, 1.9],

  // 8 · cuatro pilares (trazo navy, tamaño grande)
  s8_gobernanza:     [lu.LuGlobe,       NAVY, 1.35],
  s8_servicios:      [lu.LuWrench,      NAVY, 1.35],
  s8_proyectos:      [lu.LuLayers,      NAVY, 1.35],
  s8_alianzas:       [lu.LuHandshake,   NAVY, 1.35],

  // 11 · impacto ciudadano
  s11_salud:         [lu.LuHeartPulse,        W, 1.6],
  s11_movilidad:     [lu.LuCarFront,          W, 1.6],
  s11_seguridad:     [lu.LuShieldCheck,       W, 1.6],
  s11_alumbrado:     [lu.LuLightbulb,         W, 1.6],
  s11_ambiente:      [lu.LuLeaf,              W, 1.6],
  s11_digital:       [lu.LuMonitorSmartphone, W, 1.6],

  // viñetas y adornos
  check_green:       [lu.LuCircleCheck, GREEN, 2.0],
  quote_shield:      [lu.LuShieldCheck, W,     1.9],
};

// El original sí usa disco macizo con check blanco en las láminas 4 y 7.
const SOLID = { check_blue: [fa.FaCircleCheck, BLUE] };

(async () => {
  for (const [name, [Comp, color, sw]] of Object.entries(ICONS)) {
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(Comp, { color: "#" + color, size: 512, strokeWidth: sw })
    );
    await sharp(Buffer.from(svg), { density: 600 })
      .resize(512, 512, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
      .png().toFile(path.join(OUT, name + ".png"));
  }
  for (const [name, [Comp, color]] of Object.entries(SOLID)) {
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(Comp, { color: "#" + color, size: 512 })
    );
    await sharp(Buffer.from(svg), { density: 600 })
      .resize(512, 512, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
      .png().toFile(path.join(OUT, name + ".png"));
  }
  console.log("Iconos:", Object.keys(ICONS).length + Object.keys(SOLID).length);
})();
