const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");
const fs = require("fs");
const path = require("path");

const OUT = path.join(__dirname, "icons");
fs.mkdirSync(OUT, { recursive: true });

const W = "FFFFFF";
const NAVY = "0B2C5E";
const BLUE = "1668B8";
const GREEN = "2E9E5B";

// name -> [component, hexColor]
const ICONS = {
  // slide 1 – franja inferior de la portada
  s1_datos:            [fa.FaDatabase,           W],
  s1_interop:          [fa.FaCircleNodes,        W],
  s1_estandares:       [fa.FaShieldHalved,       W],
  s1_conectividad:     [fa.FaWifi,               W],
  s1_sostenibilidad:   [fa.FaLeaf,               W],

  // slide 2 – tarjeta flotante (blanco sobre círculo azul)
  s2_servicios:        [fa.FaUsersGear,          W],
  s2_decisiones:       [fa.FaDiagramProject,     W],
  s2_inversion:        [fa.FaBuildingColumns,    W],
  s2_ciudades:         [fa.FaCity,               W],

  // slide 2 – banda de KPIs
  s2_kpi_municipios:   [fa.FaSitemap,            W],
  s2_kpi_ciudadanos:   [fa.FaUsers,              W],
  s2_kpi_proyectos:    [fa.FaClipboardCheck,     W],
  s2_kpi_decision:     [fa.FaLightbulb,          W],

  // slide 5 – tres tarjetas
  s5_naturaleza:       [fa.FaBuildingColumns,    W],
  s5_tecnico:          [fa.FaArrowRightArrowLeft, W],
  s5_valor:            [fa.FaCheck,              W],

  // slide 6 – cómo opera OASC
  s6_gobernanza:       [fa.FaBuildingColumns,    W],
  s6_servicios:        [fa.FaBookOpen,           W],
  s6_proyectos:        [fa.FaDiagramProject,     W],
  s6_itu:              [fa.FaHandshake,          W],

  // slide 8 – cuatro pilares (línea, navy sobre claro)
  s8_gobernanza:       [fa.FaEarthAmericas,      NAVY],
  s8_servicios:        [fa.FaScrewdriverWrench,  NAVY],
  s8_proyectos:        [fa.FaLayerGroup,         NAVY],
  s8_alianzas:         [fa.FaHandshakeAngle,     NAVY],

  // slide 11 – impacto ciudadanía (blanco sobre navy)
  s11_salud:           [fa.FaHeart,              W],
  s11_movilidad:       [fa.FaCarSide,            W],
  s11_seguridad:       [fa.FaShieldHalved,       W],
  s11_alumbrado:       [fa.FaLightbulb,          W],
  s11_ambiente:        [fa.FaLeaf,               W],
  s11_digital:         [fa.FaMobileScreenButton, W],

  // vietas
  check_blue:          [fa.FaCircleCheck,        BLUE],
  check_green:         [fa.FaCircleCheck,        GREEN],
  quote_shield:        [fa.FaShieldHalved,       BLUE],
};

(async () => {
  for (const [name, [Comp, color]] of Object.entries(ICONS)) {
    const svg = ReactDOMServer.renderToStaticMarkup(
      React.createElement(Comp, { color: "#" + color, size: 512 })
    );
    await sharp(Buffer.from(svg), { density: 600 })
      .resize(512, 512, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
      .png()
      .toFile(path.join(OUT, name + ".png"));
  }
  console.log("Iconos generados:", Object.keys(ICONS).length);
})();
